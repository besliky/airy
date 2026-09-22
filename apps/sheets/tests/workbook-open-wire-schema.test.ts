/**
 * The open path parses the sidecar's open reply through a STRICT schema:
 * openWorkbookSession runs sidecarOpenResultSchema (workbookFileSchema with
 * two keys omitted — omit keeps the strictness) over the wire result, then
 * workbookFileSchema over the same object again for the renderer contract.
 * An additive sidecar reply field that misses this whitelist fails EVERY
 * workbook open with a ZodError unrecognized_keys — exactly the PR #123
 * sheets e2e regression, where the rebuilt sidecar started sending rawBytes
 * (BUG-1305) and all sheets tests died on an empty name box while every
 * unit test that mocks the protocol stayed green. This file pins the
 * tolerance: sidecar wire fields are whitelisted here before the binary
 * that sends them ships.
 */
import { describe, expect, it } from 'vitest'

import { workbookFileSchema } from '../src/shared/desktop-api'

// mirrors sidecarOpenResultSchema in sheets-main (not imported: the module
// pulls in Electron); omit() keeps the strict key checking that regressed
const sidecarOpenResultSchema = workbookFileSchema.omit({
  sha256: true,
  readOnly: true,
})

const baseSheet = {
  id: 'sheet-1',
  name: 'Sheet1',
  rowCount: 10,
  columnCount: 5,
  columnWidths: [],
  defaultRowHeight: null,
  defaultColumnWidth: null,
  freeze: null,
  hidden: false,
  tabColor: null,
  showGridLines: true,
  tables: [],
  comments: [],
  pivotRanges: [],
}

/** the sidecar open reply as the wire carries it (no sha256/readOnly yet) */
function sidecarOpenResult(rawBytes?: number) {
  return {
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    name: 'book.xlsx',
    entryCount: 3,
    sheets: [baseSheet],
    activeTab: 0,
    styles: [],
    dxfStyles: [],
    visuals: [],
    definedNames: [],
    ...(rawBytes === undefined ? {} : { rawBytes }),
  }
}

describe('workbook open wire schema vs additive sidecar fields (BUG-1305 regression)', () => {
  it('accepts the open reply carrying the new rawBytes field', () => {
    // the exact CI failure shape: the rebuilt sidecar answers every open
    // with rawBytes and the strict parse refused it (unrecognized_keys)
    const opened = sidecarOpenResultSchema.parse(sidecarOpenResult(53248))
    expect(opened.rawBytes).toBe(53248)
  })

  it('still accepts a reply without rawBytes (pre-BUG-1305 sidecar binaries)', () => {
    const opened = sidecarOpenResultSchema.parse(sidecarOpenResult())
    expect(opened.rawBytes).toBeUndefined()
  })

  it('keeps accepting rawBytes through the renderer-facing reparse', () => {
    // openWorkbookSession parses the SAME object a second time after
    // spreading it into the WorkbookFile — the whitelist must hold there too
    const workbook = workbookFileSchema.parse({
      ...sidecarOpenResult(53248),
      path: '/tmp/book.xlsx',
      sha256: 'a'.repeat(64),
      readOnly: false,
    })
    expect(workbook.rawBytes).toBe(53248)
  })

  it('rejects a malformed rawBytes instead of trusting it', () => {
    expect(() => sidecarOpenResultSchema.parse(sidecarOpenResult(-1))).toThrow()
    expect(() => sidecarOpenResultSchema.parse(sidecarOpenResult(1.5))).toThrow()
  })
})
