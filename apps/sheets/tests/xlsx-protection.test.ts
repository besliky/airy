import { describe, expect, it } from 'vitest'

import {
  applyProtectedRanges,
  applySheetProtection,
  applyWorkbookProtection,
  SheetProtectionError,
  type SheetProtectionSpec,
} from '../src/gateway/xlsx-protection'
import { excelLegacyPasswordHash } from '../src/shared/legacy-password'

const BARE = '<worksheet><sheetData/><autoFilter ref="A1:C4"/></worksheet>'

describe('applySheetProtection', () => {
  it('inserts the element after sheetData with Excel defaults', () => {
    expect(applySheetProtection(BARE, { protected: true })).toBe(
      '<worksheet><sheetData/><sheetProtection sheet="1" objects="1" scenarios="1"/>' +
        '<autoFilter ref="A1:C4"/></worksheet>',
    )
  })

  it('removes any element (password forms included) and is a no-op without one', () => {
    const protectedXml = applySheetProtection(BARE, { protected: true })
    expect(applySheetProtection(protectedXml, { protected: false })).toBe(BARE)
    expect(applySheetProtection(BARE, { protected: false })).toBe(BARE)
    // The renderer verifies the password before recording the unprotect, so
    // the gateway may strip password-bearing elements unconditionally.
    for (const attrs of [
      'sheet="1" password="83AF"',
      'sheet="1" algorithmName="SHA-512" hashValue="x" saltValue="y" spinCount="100000"',
    ]) {
      const xml = `<worksheet><sheetData/><sheetProtection ${attrs}/></worksheet>`
      expect(applySheetProtection(xml, { protected: false })).toBe(
        '<worksheet><sheetData/></worksheet>',
      )
    }
  })

  it('rebuilds an existing element from the full spec', () => {
    const xml =
      '<worksheet><sheetData/>' +
      '<sheetProtection sheet="0" formatCells="0" insertRows="0"/></worksheet>'
    expect(
      applySheetProtection(xml, {
        protected: true,
        attributes: { formatCells: false, insertRows: false },
      }),
    ).toBe(
      '<worksheet><sheetData/>' +
        '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" insertRows="0"/>' +
        '</worksheet>',
    )
  })

  it('writes the legacy password hash and raw-polarity attributes', () => {
    const spec: SheetProtectionSpec = {
      protected: true,
      passwordHash: excelLegacyPasswordHash('secret'),
      // Excel dialog "allowed" checkboxes map to attribute "0" here.
      attributes: {
        selectLockedCells: false,
        selectUnlockedCells: false,
        formatCells: false,
        insertRows: false,
        deleteRows: true,
        sort: false,
        autoFilter: false,
      },
    }
    const out = applySheetProtection(BARE, spec)
    expect(out).toContain('password="DAA7"')
    expect(out).toContain('selectLockedCells="0"')
    expect(out).toContain('formatCells="0"')
    expect(out).toContain('deleteRows="1"')
    expect(out).toContain('sort="0"')
    expect(out).not.toContain('insertColumns')
    // Unmodeled attributes are not preserved: the save sends the full spec.
    const previous =
      '<worksheet><sheetData/><sheetProtection sheet="1" pivotTables="0"/></worksheet>'
    expect(applySheetProtection(previous, { protected: true })).not.toContain('pivotTables')
  })

  it('handles the paired-tag form', () => {
    const xml = '<worksheet><sheetData/><sheetProtection sheet="1"></sheetProtection></worksheet>'
    expect(applySheetProtection(xml, { protected: false })).toBe(
      '<worksheet><sheetData/></worksheet>',
    )
  })
})

const WORKBOOK =
  '<workbook><workbookPr/><bookViews><workbookView/></bookViews>' +
  '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>'

describe('applyWorkbookProtection', () => {
  it('inserts lockStructure before bookViews', () => {
    expect(applyWorkbookProtection(WORKBOOK, true)).toBe(
      '<workbook><workbookPr/><workbookProtection lockStructure="1"/>' +
        '<bookViews><workbookView/></bookViews>' +
        '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    )
  })

  it('round-trips and drops an emptied element', () => {
    const locked = applyWorkbookProtection(WORKBOOK, true)
    expect(applyWorkbookProtection(locked, false)).toBe(WORKBOOK)
    expect(applyWorkbookProtection(WORKBOOK, false)).toBe(WORKBOOK)
  })

  it('keeps unrelated protection attributes when unlocking', () => {
    const xml = WORKBOOK.replace(
      '<bookViews>',
      '<workbookProtection lockStructure="1" lockWindows="1"/><bookViews>',
    )
    expect(applyWorkbookProtection(xml, false)).toContain('<workbookProtection lockWindows="1"/>')
  })

  it('tolerates newline/tab attribute separators', () => {
    const xml = WORKBOOK.replace(
      '<bookViews>',
      '<workbookProtection\n\tlockStructure="true"/><bookViews>',
    )
    expect(applyWorkbookProtection(xml, false)).toBe(WORKBOOK)
    const relocked = applyWorkbookProtection(
      WORKBOOK.replace('<bookViews>', '<workbookProtection\n\tlockStructure="0"/><bookViews>'),
      true,
    )
    expect(relocked).toContain('lockStructure="1"')
  })

  it('fails closed when unlocking a password-protected structure', () => {
    for (const attrs of [
      'lockStructure="1" workbookPassword="83AF"',
      'lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="x"',
    ]) {
      const xml = WORKBOOK.replace('<bookViews>', `<workbookProtection ${attrs}/><bookViews>`)
      expect(() => applyWorkbookProtection(xml, false)).toThrow(SheetProtectionError)
    }
  })
})

describe('applyProtectedRanges', () => {
  it('inserts after sheetProtection and escapes attributes', () => {
    const xml =
      '<worksheet><sheetData/><sheetProtection sheet="1"/>' +
      '<autoFilter ref="A1:C4"/></worksheet>'
    expect(applyProtectedRanges(xml, [{ name: 'R&D "range"', sqref: 'A1:B4' }])).toBe(
      '<worksheet><sheetData/><sheetProtection sheet="1"/>' +
        '<protectedRanges><protectedRange sqref="A1:B4" name="R&amp;D &quot;range&quot;"/>' +
        '</protectedRanges><autoFilter ref="A1:C4"/></worksheet>',
    )
  })

  it('replaces the existing set and removes it when emptied', () => {
    const xml =
      '<worksheet><sheetData/>' +
      '<protectedRanges><protectedRange sqref="C1" name="Old"/></protectedRanges></worksheet>'
    expect(applyProtectedRanges(xml, [{ name: 'New', sqref: 'D2:D9' }])).toBe(
      '<worksheet><sheetData/>' +
        '<protectedRanges><protectedRange sqref="D2:D9" name="New"/></protectedRanges>' +
        '</worksheet>',
    )
    expect(applyProtectedRanges(xml, [])).toBe('<worksheet><sheetData/></worksheet>')
  })

  it('fails closed on password- or permission-protected ranges', () => {
    for (const attr of ['password="83AF"', 'securityDescriptor="D:(A;;CC;;;S-1-5-21-1)"']) {
      const xml =
        '<worksheet><sheetData/><protectedRanges>' +
        `<protectedRange sqref="C1" name="Locked" ${attr}/></protectedRanges></worksheet>`
      expect(() => applyProtectedRanges(xml, [])).toThrow(SheetProtectionError)
    }
  })
})
