import { describe, expect, it } from 'vitest'

import { parseSidecarReadResult } from '../src/main/sidecar-read-validation'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'

/**
 * BUG-1711 (sheets-core audit SC-01): the sidecar has always serialized a
 * legacy `password=` hash under the attribute's own wire name (`password`),
 * while the read-path schema only accepted `passwordHash` — so a sheet
 * protected the way Excel's legacy writers (and openpyxl) record it failed
 * EVERY read_range with an unrecognized_keys error, and the raw zod JSON
 * drenched the status bar. The schema now normalizes the alias before its
 * strict parse; these tests pin the whole read-path contract: both wire
 * forms parse to one canonical shape, and anything invalid surfaces as the
 * bridge's one-sentence human error, never a zod dump.
 */

/// A complete read_range reply shaped the way sheets-main receives it from
/// the sidecar; only the sheetProtection payload varies per case.
function rangeResult(sheetProtection: unknown): Record<string, unknown> {
  return {
    cells: [{ row: 1, column: 1, value: 22, styleIndex: 0 }],
    rows: [],
    merges: [],
    hyperlinks: [],
    conditionalRules: [],
    autoFilter: null,
    autoFilterColumns: [],
    dataValidations: [],
    sheetProtection,
    rowBreaks: [],
    colBreaks: [],
    protectedRanges: [],
    pageSetup: null,
    indexedThroughRow: 3,
    indexingComplete: true,
  }
}

/// What the Rust sidecar (SheetProtectionInfo, camelCase serde) emits for
/// `<sheetProtection sheet="1" password="83AF" insertRows="0"/>`: the hash
/// rides the OOXML attribute's own name, modeled attributes keep their raw
/// polarity (false = action allowed), absent attributes stay absent.
const SIDECAR_LEGACY_FORM = {
  protected: true,
  hasPassword: true,
  password: '83AF',
  insertRows: false,
}

/// The canonical form every TS consumer (preload passthrough, renderer
/// state, edit journal, save path) reads.
const CANONICAL_FORM = {
  protected: true,
  hasPassword: true,
  passwordHash: '83AF',
  insertRows: false,
}

const RANGE_RESPONSE_FAILURE = 'Invalid workbook range response.'

function expectHumanFailure(payload: unknown): Error {
  let caught: unknown
  try {
    parseSidecarReadResult(workbookRangeResultSchema, rangeResult(payload), RANGE_RESPONSE_FAILURE)
  } catch (error) {
    caught = error
  }
  expect(caught, 'the payload must be rejected').toBeInstanceOf(Error)
  const error = caught as Error
  // Exactly the one-sentence bridge message — no zod issues, no JSON dump.
  expect(error.message).toBe(RANGE_RESPONSE_FAILURE)
  return error
}

describe('read-path sheetProtection wire contract (BUG-1711)', () => {
  it('accepts the sidecar legacy form (password) and delivers the canonical passwordHash', () => {
    const parsed = workbookRangeResultSchema.parse(rangeResult(SIDECAR_LEGACY_FORM))
    // Flags and hash arrive; the alias key is folded away.
    expect(parsed.sheetProtection).toEqual(CANONICAL_FORM)
    expect(parsed.sheetProtection).not.toHaveProperty('password')
  })

  it('accepts the canonical form (passwordHash) unchanged', () => {
    const parsed = workbookRangeResultSchema.parse(rangeResult(CANONICAL_FORM))
    expect(parsed.sheetProtection).toEqual(CANONICAL_FORM)
  })

  it('accepts a modern-hash sheet: hasPassword without any hash field stays fail-closed', () => {
    const parsed = workbookRangeResultSchema.parse(
      rangeResult({ protected: true, hasPassword: true }),
    )
    expect(parsed.sheetProtection).toEqual({ protected: true, hasPassword: true })
  })

  it('an unprotected sheet (null protection) parses as before', () => {
    const parsed = workbookRangeResultSchema.parse(rangeResult(null))
    expect(parsed.sheetProtection).toBeNull()
  })

  it('keeps the strict object strict: an unknown key still rejects with the human error', () => {
    expectHumanFailure({ ...CANONICAL_FORM, hashValue: 'x' })
  })

  it('rejects the ambiguous both-keys form with the human error', () => {
    expectHumanFailure({ ...CANONICAL_FORM, password: '83AF' })
  })

  it('rejects a non-hex hash under either name with the human error', () => {
    expectHumanFailure({ ...CANONICAL_FORM, passwordHash: 'xyz!' })
    expectHumanFailure({ ...SIDECAR_LEGACY_FORM, password: 'secret' })
  })

  it('rejects a wrong-typed hash with the human error', () => {
    expectHumanFailure({ ...SIDECAR_LEGACY_FORM, password: 83 })
  })
})

describe('parseSidecarReadResult error convention (BUG-1711)', () => {
  it('passes valid data through untouched and throws only the given sentence on failure', () => {
    const valid = rangeResult({ protected: false, hasPassword: false })
    expect(
      parseSidecarReadResult(workbookRangeResultSchema, valid, RANGE_RESPONSE_FAILURE),
    ).toEqual(workbookRangeResultSchema.parse(valid))
    // A non-object sidecar reply takes the same human path.
    expect(() =>
      parseSidecarReadResult(workbookRangeResultSchema, 42, RANGE_RESPONSE_FAILURE),
    ).toThrow(RANGE_RESPONSE_FAILURE)
  })
})
