import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { createBufferEntrySource, planCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetThreadedCommentState } from '../src/gateway/xlsx-gateway'
import {
  applyThreadedCommentStates,
  buildPersonsXml,
  buildThreadedCommentsXml,
} from '../src/gateway/xlsx-threaded-comments'
import { buildEditFixture } from './fixture-builder'

/// planCellEditsToXlsx is positional; only the threaded-comment states vary
/// across these tests (last parameter).
async function planThreads(states: SheetThreadedCommentState[], source?: Uint8Array) {
  const entrySource = await createBufferEntrySource(source ?? (await buildEditFixture()))
  return planCellEditsToXlsx(
    entrySource,
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
    states,
  )
}

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
            author: 'Reviewer <lead>',
            text: 'Tax inclusive <confirm>',
            createdAt: '2026-09-23T10:00:00.000Z',
          },
          {
            id: '{22222222-2222-2222-2222-222222222222}',
            personId: '{aaaaaaaa-2222-2222-2222-222222222222}',
            author: 'Airy user',
            text: 'Done — see cell "C2"',
            createdAt: '2026-09-23T11:30:00.000Z',
          },
        ],
      },
      {
        id: '{33333333-3333-3333-3333-333333333333}',
        row: 4,
        column: 0,
        resolved: false,
        messages: [
          {
            id: '{33333333-3333-3333-3333-333333333333}',
            personId: '{aaaaaaaa-2222-2222-2222-222222222222}',
            author: 'Airy user',
            text: 'second thread',
            createdAt: '2026-09-23T12:00:00.000Z',
          },
        ],
      },
    ],
  },
]

/// A fixture that already carries a threadedComments part (as Excel writes
/// it), so removal and rewrite paths have something to work on.
async function buildThreadedFixture(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  zip.file(
    'xl/threadedComments/threadedComment1.xml',
    buildThreadedCommentsXml([
      {
        id: '{44444444-4444-4444-4444-444444444444}',
        row: 0,
        column: 0,
        resolved: false,
        messages: [
          {
            id: '{44444444-4444-4444-4444-444444444444}',
            personId: '{aaaaaaaa-3333-3333-3333-333333333333}',
            author: 'Original',
            text: 'pre-existing',
            createdAt: '2026-01-01T00:00:00.000Z',
          },
        ],
      },
    ]),
  )
  zip.file(
    'xl/persons/person.xml',
    buildPersonsXml([
      {
        id: '{aaaaaaaa-3333-3333-3333-333333333333}',
        displayName: 'Original',
        userId: 'user-3333',
        providerId: 'AD',
      },
    ]),
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>`,
  )
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedComments+xml"/><Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('threaded comment serialization', () => {
  it('writes the threadedComments part with refs, replies, done, and escaped text', async () => {
    const plan = await planThreads(STATES)
    const part = [...plan.added.entries()].find(([path]) =>
      /xl\/threadedComments\/threadedComment\d+\.xml/.test(path),
    )
    expect(part).toBeDefined()
    const xml = part![1]
    // Root: anchored, carries the resolution flag, escapes XML specials.
    expect(xml).toContain(
      '<threadedComment dT="2026-09-23T10:00:00.000Z" ref="C2" personId="{aaaaaaaa-1111-1111-1111-111111111111}" id="{11111111-1111-1111-1111-111111111111}" done="1">',
    )
    expect(xml).toContain('Tax inclusive &lt;confirm&gt;')
    // Reply: no ref, points at its root through parentId.
    expect(xml).toContain(
      '<threadedComment dT="2026-09-23T11:30:00.000Z" personId="{aaaaaaaa-2222-2222-2222-222222222222}" id="{22222222-2222-2222-2222-222222222222}" parentId="{11111111-1111-1111-1111-111111111111}">',
    )
    expect(xml).toContain('Done — see cell &quot;C2&quot;')
    expect(xml).toContain('ref="A5"')
  })

  it('writes the persons directory and registers rels and content types', async () => {
    const plan = await planThreads(STATES)
    const persons = [...plan.added.entries()].find(([path]) => path === 'xl/persons/person.xml')
    expect(persons).toBeDefined()
    expect(persons![1]).toContain(
      '<person id="{aaaaaaaa-1111-1111-1111-111111111111}" displayName="Reviewer &lt;lead&gt;" providerId="None"/>',
    )
    const workbookRels = plan.replaced.get('xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain('office/2017/10/relationships/person')
    const sheetRels = plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain('office/2017/10/relationships/threadedComment')
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('vnd.ms-excel.threadedComments+xml')
    expect(contentTypes).toContain('vnd.ms-excel.person+xml')
  })

  it('removes the sheet part and persons directory when the thread set empties', async () => {
    const plan = await planThreads(
      [{ sheetName: 'Data', threads: [] }],
      await buildThreadedFixture(),
    )
    expect(plan.removedEntries).toContain('xl/threadedComments/threadedComment1.xml')
    expect(plan.removedEntries).toContain('xl/persons/person.xml')
    const sheetRels = plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).not.toContain('threadedComment')
    const workbookRels = plan.replaced.get('xl/_rels/workbook.xml.rels')
    expect(workbookRels).not.toContain('relationships/person')
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).not.toContain('threadedComments')
    expect(contentTypes).not.toContain('vnd.ms-excel.person+xml')
  })

  it('rewrites an existing part in place and preserves unknown persons', async () => {
    const plan = await planThreads(STATES, await buildThreadedFixture())
    const rewritten = [...plan.replaced.entries()].find(([path]) =>
      /threadedComment\d+\.xml/.test(path),
    )
    expect(rewritten).toBeDefined()
    expect(rewritten![0]).toBe('xl/threadedComments/threadedComment1.xml')
    const persons = plan.replaced.get('xl/persons/person.xml')
    expect(persons).toContain('displayName="Original"')
    // The untouched sheet threads' author survives the merge…
    expect(persons).toContain('user-3333')
    // …and this save's mentions are present alongside.
    expect(persons).toContain('displayName="Reviewer &lt;lead&gt;"')
  })

  it('is a no-op when a sheet without threads is not covered by the save', async () => {
    const plan = await planThreads([])
    expect([...plan.added.keys()].some((path) => path.includes('threadedComment'))).toBe(false)
    expect(plan.removedEntries).toHaveLength(0)
  })
})

describe('threaded comment persons merge (direct package apply)', () => {
  it('keeps existing person entries and appends the save mentions', async () => {
    const added = new Map<string, string>()
    const replaced = new Map<string, string>()
    const removed = new Set<string>()
    const paths = new Set<string>([
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/_rels/sheet1.xml.rels',
      'xl/persons/person.xml',
      'xl/threadedComments/threadedComment7.xml',
      '[Content_Types].xml',
    ])
    const pkg = {
      paths: async () => [...paths],
      has: async (path: string) => paths.has(path) && !removed.has(path),
      readText: async (path: string) => {
        if (path === 'xl/persons/person.xml') {
          return buildPersonsXml([
            { id: '{p-keep}', displayName: 'Kept Author', userId: 'u1', providerId: 'AD' },
          ])
        }
        if (path === 'xl/_rels/workbook.xml.rels') {
          return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>`
        }
        if (path === 'xl/worksheets/_rels/sheet1.xml.rels') {
          return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment7.xml"/></Relationships>`
        }
        return '<Types/>'
      },
      write: (path: string, content: string) => {
        replaced.set(path, content)
      },
      add: (path: string, content: string) => {
        added.set(path, content)
      },
      remove: (path: string) => {
        removed.add(path)
      },
    }
    await applyThreadedCommentStates(
      pkg,
      [
        {
          worksheetPath: 'xl/worksheets/sheet1.xml',
          threads: [
            {
              id: '{t1}',
              row: 0,
              column: 1,
              resolved: false,
              messages: [
                {
                  id: '{t1}',
                  personId: '{p-new}',
                  author: 'New Author',
                  text: 'hello',
                  createdAt: '2026-09-23T09:00:00.000Z',
                },
              ],
            },
          ],
        },
      ],
      ['xl/worksheets/sheet1.xml'],
      new Set<string>(),
    )
    const persons = replaced.get('xl/persons/person.xml') ?? added.get('xl/persons/person.xml')
    expect(persons).toContain('Kept Author')
    expect(persons).toContain('New Author')
    expect(persons!.indexOf('Kept Author')).toBeLessThan(persons!.indexOf('New Author'))
    // The rewritten threadedComments part replaces the stale one in place.
    expect(added.has('xl/threadedComments/threadedComment7.xml')).toBe(false)
    expect(replaced.get('xl/threadedComments/threadedComment7.xml')).toContain('ref="B1"')
  })
})
