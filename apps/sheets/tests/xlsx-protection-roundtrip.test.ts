import { execFileSync } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import JSZip from 'jszip'
import { afterAll, describe, expect, it } from 'vitest'

import { applySheetProtection } from '../src/gateway/xlsx-protection'
import { excelLegacyPasswordHash } from '../src/shared/legacy-password'

/**
 * PAR-204 round-trip for the <sheetProtection> element our save writes, in
 * two layers:
 *   - Pure JS (runs everywhere, CI included): the protected sheet is saved
 *     into a workbook package, the package is unzipped again and its
 *     worksheet XML is asserted directly — the protection flag, the legacy
 *     password hash (DAA7 for 'secret') and each attribute's raw OOXML
 *     polarity. This is the coverage GitHub runners get: they ship neither
 *     python nor openpyxl.
 *   - openpyxl (local bonus): the same package is reopened by a third-party
 *     OOXML reader, the stand-in for Excel itself. Skipped with an honest
 *     label wherever `python3 -c "import openpyxl"` fails.
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

/// One availability probe for the openpyxl layer: python3 must exist AND be
/// able to import openpyxl (ModuleNotFoundError on GitHub runners => false,
/// as does a missing python3 entirely).
const openpyxlAvailable = ((): boolean => {
  try {
    execFileSync('python3', ['-c', 'import openpyxl'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
})()

let dirPromise: Promise<string> | null = null
const workDir = (): Promise<string> =>
  (dirPromise ??= mkdtemp(join(tmpdir(), 'par204-protection-')))
afterAll(async () => {
  if (dirPromise) await rm(await dirPromise, { recursive: true, force: true })
})

/// Reopens the saved package the way a third-party reader would: unzip it
/// and return the worksheet part.
async function readSavedSheetXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const entry = zip.file('xl/worksheets/sheet1.xml')
  if (!entry) throw new Error('The saved package has no xl/worksheets/sheet1.xml part.')
  return entry.async('string')
}

const protectionElement = (xml: string): string => /<sheetProtection\b[^>]*\/>/.exec(xml)?.[0] ?? ''

function attr(element: string, name: string): string | null {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(element)?.[1] ?? null
}

describe('sheetProtection saved package (pure JS — the CI coverage, no python)', () => {
  it('our protected save writes the flag, legacy hash and raw attribute polarity', async () => {
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
    const xml = await readSavedSheetXml(
      await writeWorkbook(applySheetProtection(WORKSHEET, spec), dir, 'protected.xlsx'),
    )
    const element = protectionElement(xml)
    expect(element).not.toBe('')
    expect(attr(element, 'sheet')).toBe('1')
    // Excel's writer header: the structure itself stays lockable.
    expect(attr(element, 'objects')).toBe('1')
    expect(attr(element, 'scenarios')).toBe('1')
    // The legacy hash a third-party reader maps back to ws.protection.password.
    expect(attr(element, 'password')).toBe('DAA7')
    // Polarity is written explicitly and mirrors the dialog state:
    // "0" = allowed, "1" = prevented.
    for (const [name, prevented] of Object.entries(spec.attributes)) {
      expect(attr(element, name)).toBe(prevented ? '1' : '0')
    }
    // The sheet data itself survives the rewrite untouched.
    expect(xml).toContain('<v>1</v>')
  })

  it('our default protect is Excel Minimal: flag on, no hash, schema-default attributes', async () => {
    const dir = await workDir()
    const xml = await readSavedSheetXml(
      await writeWorkbook(
        applySheetProtection(WORKSHEET, { protected: true }),
        dir,
        'default-protected.xlsx',
      ),
    )
    const element = protectionElement(xml)
    expect(element).not.toBe('')
    expect(attr(element, 'sheet')).toBe('1')
    // Absent password= is what a third-party reader reports as "no hash".
    expect(attr(element, 'password')).toBeNull()
    // Modeled attributes are absent, so the schema defaults decide: every
    // action prevented, selecting allowed — Excel Minimal.
    expect(attr(element, 'selectLockedCells')).toBeNull()
    expect(attr(element, 'selectUnlockedCells')).toBeNull()
    expect(attr(element, 'formatCells')).toBeNull()
    expect(attr(element, 'insertRows')).toBeNull()
    expect(attr(element, 'sort')).toBeNull()
    expect(attr(element, 'autoFilter')).toBeNull()
  })

  it('unprotect removes the element from the saved package entirely', async () => {
    const dir = await workDir()
    const protectedXml = applySheetProtection(WORKSHEET, {
      protected: true,
      passwordHash: 'DAA7',
    })
    const xml = await readSavedSheetXml(
      await writeWorkbook(
        applySheetProtection(protectedXml, { protected: false }),
        dir,
        'unprotected.xlsx',
      ),
    )
    expect(xml).not.toContain('sheetProtection')
    expect(xml).toContain('<v>1</v>')
  })
})

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

describe.skipIf(!openpyxlAvailable)(
  'sheetProtection openpyxl round-trip (local bonus — skipped honestly when python3+openpyxl is unavailable)',
  () => {
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
  },
)
