import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { SheetThreadedCommentState } from '../src/gateway/xlsx-gateway'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { buildEditFixture } from './fixture-builder'

/// Round-trip verification for threaded comments: a save that carries the
/// new part set must (a) load cleanly in openpyxl (strict zip + rels + XML
/// consumers must not choke on the new parts) and (b) re-open in our own
/// sidecar with the threads resolved back into anchors/authors/replies.
/// External consumers are skipped gracefully where unavailable.

const PYTHON = ['/usr/bin/python3', '/usr/local/bin/python3'].find((path) => existsSync(path))
const HAS_OPENPYXL =
  PYTHON !== undefined &&
  (() => {
    try {
      execFileSync(PYTHON, ['-c', 'import openpyxl'], { stdio: 'ignore' })
      return true
    } catch {
      return false
    }
  })()

const STATES: SheetThreadedCommentState[] = [
  {
    sheetName: 'Data',
    threads: [
      {
        id: '{11111111-1111-1111-1111-111111111111}',
        row: 1,
        column: 2,
        resolved: true,
        messages: [
          {
            id: '{11111111-1111-1111-1111-111111111111}',
            personId: '{aaaaaaaa-1111-1111-1111-111111111111}',
            author: 'Reviewer',
            text: 'Tax inclusive',
            createdAt: '2026-09-23T10:00:00.000Z',
          },
          {
            id: '{22222222-2222-2222-2222-222222222222}',
            personId: '{aaaaaaaa-1111-1111-1111-111111111111}',
            author: 'Reviewer',
            text: 'Confirmed',
            createdAt: '2026-09-23T11:30:00.000Z',
          },
        ],
      },
    ],
  },
]

async function saveThreadedWorkbook(): Promise<Buffer> {
  const source = await createBufferEntrySource(await buildEditFixture())
  const plan = await planCellEditsToXlsx(
    source,
    [], // edits
    [], // structuralOps
    [], // chartEdits
    undefined, // sheetPlan
    [], // filterStates
    [], // hyperlinkEdits
    [], // cfStates
    [], // dvStates
    [], // sheetProtections
    null, // definedNamesState
    [], // visualAdditions
    [], // pageSetupStates
    [], // noteStates
    [], // tableAdditions
    [], // pivotAdditions
    [], // pivotCacheRefreshPaths
    [], // pivotRefreshUpdates
    [], // visualEdits
    [], // sparklineAdditions
    [], // formulaValues
    null, // themeState
    null, // workbookProtectionState
    [], // protectedRangeStates
    [], // bulkConstantFills
    STATES,
  )
  return (await assembleWithJsZip(await buildEditFixture(), plan)).buffer
}

describe('threaded comments round-trip', () => {
  it('the saved package carries the parts Excel expects', async () => {
    const buffer = await saveThreadedWorkbook()
    const zip = await JSZip.loadAsync(buffer)
    const threadPart = zip.file('xl/threadedComments/threadedComment1.xml')
    expect(threadPart).toBeDefined()
    const xml = await threadPart!.async('string')
    expect(xml).toContain('done="1"')
    expect(xml).toContain('ref="C2"')
    expect(xml).toContain('parentId="{11111111-1111-1111-1111-111111111111}"')
    expect(zip.file('xl/persons/person.xml')).toBeDefined()
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string')
    expect(contentTypes).toContain('application/vnd.ms-excel.threadedComments+xml')
    expect(contentTypes).toContain('application/vnd.ms-excel.person+xml')
  })

  it.skipIf(!HAS_OPENPYXL)('openpyxl loads the saved workbook without error', async () => {
    const buffer = await saveThreadedWorkbook()
    const dir = await mkdtemp(join(tmpdir(), 'airy-threads-'))
    try {
      const path = join(dir, 'threaded.xlsx')
      await writeFile(path, buffer)
      // Loading is the verification: openpyxl parses [Content_Types].xml, the
      // rels graph, and every part it knows; unknown parts must be ignored.
      execFileSync(
        PYTHON!,
        ['-c', 'import openpyxl, sys; openpyxl.load_workbook(sys.argv[1])', path],
        { stdio: 'pipe' },
      )
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it.skipIf(!existsSync(sidecarBinaryPath()))(
    'our sidecar re-opens the saved workbook with the threads resolved',
    async () => {
      const buffer = await saveThreadedWorkbook()
      const dir = await mkdtemp(join(tmpdir(), 'airy-threads-'))
      const client = new XlsxSidecarClient(sidecarBinaryPath())
      let sessionId: string | null = null
      try {
        const path = join(dir, 'threaded.xlsx')
        await writeFile(path, buffer)
        const opened = (await client.open(path)) as {
          sessionId: string
          sheets: {
            threadedComments?: {
              id: string
              row: number
              column: number
              author: string
              done: boolean
              parentId?: string
            }[]
          }[]
        }
        sessionId = opened.sessionId
        const threads = opened.sheets[0]!.threadedComments ?? []
        expect(threads).toHaveLength(2)
        expect(threads[0]).toMatchObject({
          id: '{11111111-1111-1111-1111-111111111111}',
          row: 1,
          column: 2,
          author: 'Reviewer',
          done: true,
        })
        // The reply comes back attached to its root, anchored to the same cell.
        expect(threads[1]).toMatchObject({
          id: '{22222222-2222-2222-2222-222222222222}',
          parentId: '{11111111-1111-1111-1111-111111111111}',
          row: 1,
          column: 2,
          done: false,
        })
      } finally {
        if (sessionId !== null) await client.close(sessionId).catch(() => undefined)
        client.stop()
        await rm(dir, { recursive: true, force: true })
      }
    },
  )
})

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}
