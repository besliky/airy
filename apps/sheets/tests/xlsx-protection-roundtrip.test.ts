import { execFileSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterAll, describe, expect, it } from 'vitest'

import { applySheetProtection } from '../src/gateway/xlsx-protection'
import { excelLegacyPasswordHash } from '../src/shared/legacy-password'

/**
 * PAR-204 round-trip: the sheetProtection element our save writes must be
 * recognized by a third-party OOXML reader — openpyxl is the stand-in for
 * Excel itself. Verified in both directions:
 *   - protect (with legacy password + allowed-action attributes) → openpyxl
 *     sees sheet=true, the hash, and each attribute's raw polarity;
 *   - unprotect → openpyxl sees a plain unprotected sheet.
 */

const WORKSHEET =
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

async function writeWorkbook(worksheetXml: string, dir: string, name: string): Promise<string> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', worksheetXml)
  const path = join(dir, name)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
  return path
}

/// Reads the sheet's protection state back through openpyxl as plain JSON.
function readProtection(path: string): Record<string, unknown> {
  const script = [
    'import json, sys',
    'import openpyxl',
    'wb = openpyxl.load_workbook(sys.argv[1])',
    'ws = wb.active',
    'p = ws.protection',
    'print(json.dumps({',
    '  "sheet": p.sheet,',
    '  "password": p.password,',
    '  "selectLockedCells": p.selectLockedCells,',
    '  "selectUnlockedCells": p.selectUnlockedCells,',
    '  "formatCells": p.formatCells,',
    '  "formatColumns": p.formatColumns,',
    '  "formatRows": p.formatRows,',
    '  "insertColumns": p.insertColumns,',
    '  "insertRows": p.insertRows,',
    '  "deleteColumns": p.deleteColumns,',
    '  "deleteRows": p.deleteRows,',
    '  "sort": p.sort,',
    '  "autoFilter": p.autoFilter,',
    '}))',
  ].join('\n')
  const stdout = execFileSync('python3', ['-c', script, path], { encoding: 'utf8' })
  return JSON.parse(stdout.trim()) as Record<string, unknown>
}

let dirPromise: Promise<string> | null = null
const workDir = (): Promise<string> =>
  (dirPromise ??= mkdtemp(join(tmpdir(), 'par204-protection-')))
afterAll(async () => {
  if (dirPromise) await rm(await dirPromise, { recursive: true, force: true })
})

describe('sheetProtection openpyxl round-trip', () => {
  it('our protected save is recognized: flags, hash and attribute polarity', async () => {
    const dir = await workDir()
    const spec = {
      protected: true,
      passwordHash: excelLegacyPasswordHash('secret'),
      attributes: {
        selectLockedCells: false,
        selectUnlockedCells: false,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertColumns: false,
        insertRows: false,
        deleteColumns: true,
        deleteRows: true,
        sort: false,
        autoFilter: false,
      },
    }
    const path = await writeWorkbook(applySheetProtection(WORKSHEET, spec), dir, 'protected.xlsx')
    const protection = readProtection(path)
    expect(protection.sheet).toBe(true)
    // openpyxl exposes the raw legacy hash it read from password=.
    expect(protection.password).toBe('DAA7')
    // Attribute polarity mirrors the file: false = allowed by the dialog.
    expect(protection.selectLockedCells).toBe(false)
    expect(protection.selectUnlockedCells).toBe(false)
    expect(protection.formatCells).toBe(false)
    expect(protection.formatColumns).toBe(false)
    expect(protection.formatRows).toBe(false)
    expect(protection.insertColumns).toBe(false)
    expect(protection.insertRows).toBe(false)
    // Prevented-by-default actions we did not allow read as locked.
    expect(protection.deleteColumns).toBe(true)
    expect(protection.deleteRows).toBe(true)
    expect(protection.sort).toBe(false)
    expect(protection.autoFilter).toBe(false)
  })

  it('our default protect matches Excel Minimal: prevent everything, allow selecting', async () => {
    const dir = await workDir()
    const path = await writeWorkbook(
      applySheetProtection(WORKSHEET, { protected: true }),
      dir,
      'default-protected.xlsx',
    )
    const protection = readProtection(path)
    expect(protection.sheet).toBe(true)
    // Absent password= reads back as null (no hash).
    expect(protection.password).toBeNull()
    expect(protection.formatCells).toBe(true)
    expect(protection.insertRows).toBe(true)
    expect(protection.sort).toBe(true)
    expect(protection.selectLockedCells).toBe(false)
  })

  it('unprotect removes the element entirely: openpyxl sees a plain sheet', async () => {
    const dir = await workDir()
    const protectedXml = applySheetProtection(WORKSHEET, {
      protected: true,
      passwordHash: 'DAA7',
    })
    const path = await writeWorkbook(
      applySheetProtection(protectedXml, { protected: false }),
      dir,
      'unprotected.xlsx',
    )
    const protection = readProtection(path)
    expect(protection.sheet).toBe(false)
    expect(protection.password).toBeNull()
  })
})
