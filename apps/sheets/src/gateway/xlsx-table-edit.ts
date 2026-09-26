/**
 * Editing Excel tables already stored in the file (PAR-202): Resize Table,
 * Rename Table, style tweaks, and Convert to Range.
 *
 * Like table additions this is fail-closed: coordinates are pinned (a resize
 * refuses to ride with row/column shifts on the same sheet), names are
 * validated against Excel's rules, and formula rewrites follow Excel's
 * documented semantics (rename updates references textually; Convert to
 * Range turns structured references into equivalent A1 references).
 */

import {
  renameTableInFormulaText,
  SHEET_MAX_COLUMNS,
  SHEET_MAX_ROWS,
  tableNameError,
  tableRefsToA1InFormulaText,
} from '../domain/table-refs'
import { relsPathFor, resolveRelTarget, type MutablePackage } from './xlsx-drawing-add'
import {
  areaToRef,
  collectExistingTableNames,
  parseRef,
  tablePartPaths,
  type TableArea,
} from './xlsx-table-add'

export class TableEditError extends Error {}

export interface TableStyleEdit {
  /// Built-in style name; omitted keeps the current one.
  readonly style?: string | undefined
  readonly bandedRows?: boolean | undefined
}

/// One edit against a table part reachable from `worksheetPath` by its
/// current name (the `displayName` token structured references use).
export interface SheetTableEdit {
  readonly worksheetPath: string
  readonly tableName: string
  readonly rename?: string | undefined
  /// New area; header-anchored — the start row/column must stay put.
  readonly resize?: { readonly area: TableArea } | undefined
  readonly style?: TableStyleEdit | undefined
  readonly convertToRange?: boolean | undefined
}

const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const WORKSHEET_PATTERN = /^xl\/worksheets\/(?!_rels)[^/]+\.xml$/

interface ParsedTable {
  readonly path: string
  readonly relId: string
  /// The displayName token structured references use.
  readonly name: string
  readonly area: TableArea
  readonly headerRowCount: number
  readonly totalsRowCount: number
  readonly columnNames: readonly string[]
  readonly nextColumnId: number
}

export async function applyTableEdits(
  pkg: MutablePackage,
  edits: readonly SheetTableEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (edits.length === 0) return
  for (const edit of edits) {
    const table = await findTablePart(pkg, edit.worksheetPath, edit.tableName)
    if (!table) {
      throw new TableEditError(`Table "${edit.tableName}" was not found on this sheet.`)
    }
    let xml = await pkg.readText(table.path)
    if (edit.resize) xml = await applyResize(pkg, edit, table, xml)
    if (edit.rename !== undefined && edit.rename !== table.name) {
      xml = await applyRename(pkg, edit, table, xml, touchedEntries)
    }
    if (edit.style) xml = applyStyleEdit(xml, edit.style)
    if (edit.convertToRange === true) {
      await applyConvertToRange(pkg, edit.worksheetPath, table, touchedEntries)
      continue
    }
    pkg.write(table.path, xml)
    touchedEntries.add(table.path)
  }
}

/// Locates the worksheet's table part whose name (then displayName) matches.
export async function findTablePart(
  pkg: MutablePackage,
  worksheetPath: string,
  tableName: string,
): Promise<ParsedTable | null> {
  const relsPath = relsPathFor(worksheetPath)
  if (!(await pkg.has(relsPath))) return null
  const relsXml = await pkg.readText(relsPath)
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0]
    if (!tag.includes(`Type="${TABLE_REL_TYPE}"`)) continue
    const relId = /\bId="([^"]+)"/.exec(tag)?.[1]
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1]
    if (!relId || !target) continue
    const tablePath = resolveRelTarget(worksheetPath, target)
    if (!(await pkg.has(tablePath))) continue
    const table = await parseTablePart(pkg, tablePath, relId)
    if (!table || table.name.toLowerCase() !== tableName.toLowerCase()) continue
    return table
  }
  return null
}

async function parseTablePart(
  pkg: MutablePackage,
  tablePath: string,
  relId: string,
): Promise<ParsedTable | null> {
  const xml = await pkg.readText(tablePath)
  const open = /<table\b[^>]*>/.exec(xml)?.[0]
  const ref = open === undefined ? undefined : /\bref="([^"]+)"/.exec(open)?.[1]
  if (!open || !ref) return null
  const displayName = decodeEntities(/\bdisplayName="([^"]*)"/.exec(open)?.[1] ?? '')
  const name = decodeEntities(/\bname="([^"]*)"/.exec(open)?.[1] ?? '')
  let nextColumnId = 0
  const columnNames = [...xml.matchAll(/<tableColumn\b[^>]*\/>/g)].map((column) => {
    const id = Number(/\bid="(\d+)"/.exec(column[0])?.[1] ?? '0')
    nextColumnId = Math.max(nextColumnId, id)
    return decodeEntities(/\bname="([^"]*)"/.exec(column[0])?.[1] ?? '')
  })
  return {
    path: tablePath,
    relId,
    name: displayName || name,
    area: parseRef(ref),
    headerRowCount: Number(/\bheaderRowCount="(\d+)"/.exec(open)?.[1] ?? '1'),
    totalsRowCount: Number(/\btotalsRowCount="(\d+)"/.exec(open)?.[1] ?? '0'),
    columnNames,
    nextColumnId: nextColumnId + 1,
  }
}

/// Resize: keeps the header cell anchored and rewrites the part's ref and
/// autoFilter ref plus the column list. Growing appends Excel-style
/// "ColumnN" names; shrinking drops the trailing columns. Cells are never
/// touched — shrinking just releases cells back to the sheet, and the grown
/// region stays blank for the user to fill.
async function applyResize(
  pkg: MutablePackage,
  edit: SheetTableEdit,
  table: ParsedTable,
  xml: string,
): Promise<string> {
  const area = edit.resize?.area
  if (!area) return xml
  if (area.startRow !== table.area.startRow || area.startColumn !== table.area.startColumn) {
    throw new TableEditError(
      `Table "${edit.tableName}" resizes from its header cell — the top-left corner cannot move.`,
    )
  }
  const width = area.endColumn - area.startColumn + 1
  const dataRows = area.endRow - area.startRow + 1 - table.headerRowCount - table.totalsRowCount
  if (width < 1) {
    throw new TableEditError(`Table "${edit.tableName}" must keep at least one column.`)
  }
  if (dataRows < 1) {
    throw new TableEditError(
      `Table "${edit.tableName}" needs a header row plus at least one data row.`,
    )
  }
  if (area.endRow >= SHEET_MAX_ROWS || area.endColumn >= SHEET_MAX_COLUMNS) {
    throw new TableEditError(`Table "${edit.tableName}" does not fit on the sheet.`)
  }
  await assertNoOverlapWithOtherTables(pkg, edit.tableName, table.path, area)
  await assertNoSheetConflicts(pkg, edit.worksheetPath, edit.tableName, area)

  const ref = areaToRef(area)
  let resized = xml.replace(
    /(<table\b[^>]*?\bref=")[^"]*(")/,
    (_full, prefix: string, suffix: string) => `${prefix}${ref}${suffix}`,
  )
  resized = resized.replace(
    /(<autoFilter\b[^>]*?\bref=")[^"]*(")/,
    (_full, prefix: string, suffix: string) => `${prefix}${ref}${suffix}`,
  )
  return reshapeTableColumns(resized, table, width)
}

/// Grows or trims the `<tableColumns>` list to `width` entries.
function reshapeTableColumns(xml: string, table: ParsedTable, width: number): string {
  const columnsBlock = /<tableColumns\b[^>]*>([\s\S]*?)<\/tableColumns>/.exec(xml)
  if (!columnsBlock) {
    throw new TableEditError(`Table "${table.name}" has no readable column list.`)
  }
  const elements = [...columnsBlock[1]!.matchAll(/<tableColumn\b[^>]*?\/>/g)].map(
    (column) => column[0],
  )
  if (elements.length === 0) {
    throw new TableEditError(`Table "${table.name}" has no readable column list.`)
  }
  let columns = elements
  if (width > elements.length) {
    const used = new Set(table.columnNames.map((name) => name.trim().toLowerCase()))
    let nextId = table.nextColumnId
    for (let index = elements.length; index < width; index += 1) {
      let candidate = `Column${nextId}`
      for (let suffix = 2; used.has(candidate.trim().toLowerCase()); suffix += 1) {
        candidate = `Column${nextId}_${suffix}`
      }
      used.add(candidate.trim().toLowerCase())
      columns.push(`<tableColumn id="${nextId}" name="${escapeAttribute(candidate)}"/>`)
      nextId += 1
    }
  } else if (width < elements.length) {
    columns = elements.slice(0, width)
  }
  const block = `<tableColumns count="${columns.length}">${columns.join('')}</tableColumns>`
  return xml.replace(columnsBlock[0], block)
}

/// Rename: validates Excel's table-name rules and uniqueness, updates the
/// part, and rewrites every structured reference to the old name across all
/// worksheet formulas and workbook defined names.
async function applyRename(
  pkg: MutablePackage,
  edit: SheetTableEdit,
  table: ParsedTable,
  xml: string,
  touchedEntries: Set<string>,
): Promise<string> {
  const nextName = edit.rename ?? ''
  const nameError = tableNameError(nextName)
  if (nameError) throw new TableEditError(nameError)
  const taken = await collectExistingTableNames(pkg)
  taken.delete(table.name.toLowerCase())
  if (taken.has(nextName.toLowerCase())) {
    throw new TableEditError(`Table name "${nextName}" is already taken in this workbook.`)
  }
  await assertNameNotDefined(pkg, nextName)

  const open = /<table\b[^>]*>/.exec(xml)?.[0]
  if (!open) throw new TableEditError(`Table "${edit.tableName}" part is malformed.`)
  let replaced = open.replace(/\bdisplayName="[^"]*"/, `displayName="${escapeAttribute(nextName)}"`)
  replaced = /\bname="/.test(replaced)
    ? replaced.replace(/\bname="[^"]*"/, `name="${escapeAttribute(nextName)}"`)
    : replaced.replace(/\bdisplayName="/, `name="${escapeAttribute(nextName)}" displayName="`)
  const renamed = xml.replace(open, replaced)

  await rewriteFormulasEverywhere(pkg, touchedEntries, (formula) =>
    renameTableInFormulaText(formula, table.name, nextName),
  )
  return renamed
}

/// Style edit: rewrites the tableStyleInfo style name and/or banding flag.
function applyStyleEdit(xml: string, style: TableStyleEdit): string {
  const info = /<tableStyleInfo\b[^>]*\/>/.exec(xml)?.[0]
  if (!info) {
    if (style.style === undefined && style.bandedRows === undefined) return xml
    const element =
      `<tableStyleInfo name="${escapeAttribute(style.style ?? 'TableStyleMedium2')}" ` +
      'showFirstColumn="0" showLastColumn="0" ' +
      `showRowStripes="${style.bandedRows === false ? 0 : 1}" showColumnStripes="0"/>`
    return xml.replace('</table>', `${element}</table>`)
  }
  let replaced = info
  if (style.style !== undefined) {
    replaced = /\bname="/.test(replaced)
      ? replaced.replace(/\bname="[^"]*"/, `name="${escapeAttribute(style.style)}"`)
      : replaced.replace(/\/>$/, ` name="${escapeAttribute(style.style)}"/>`)
  }
  if (style.bandedRows !== undefined) {
    replaced = /\bshowRowStripes="/.test(replaced)
      ? replaced.replace(/\bshowRowStripes="[^"]*"/, `showRowStripes="${style.bandedRows ? 1 : 0}"`)
      : replaced.replace(/\/>$/, ` showRowStripes="${style.bandedRows ? 1 : 0}"/>`)
  }
  return xml.replace(info, replaced)
}

/// Convert to Range: structured references across the workbook become their
/// equivalent A1 references, then the table part, its relationship, its
/// worksheet tablePart entry, and its content type are removed. Cell data
/// and formatting are not touched.
async function applyConvertToRange(
  pkg: MutablePackage,
  worksheetPath: string,
  table: ParsedTable,
  touchedEntries: Set<string>,
): Promise<void> {
  const sheetNames = await sheetNamesByPath(pkg)
  const sheetName = sheetNames.get(worksheetPath) ?? ''
  const geometry = {
    startRow: table.area.startRow,
    endRow: table.area.endRow,
    startColumn: table.area.startColumn,
    endColumn: table.area.endColumn,
    headerRowCount: table.headerRowCount,
    totalsRowCount: table.totalsRowCount,
  }
  await rewriteFormulasEverywhere(pkg, touchedEntries, (formula, row, formulaSheet) =>
    tableRefsToA1InFormulaText(
      formula,
      {
        name: table.name,
        geometry,
        columns: table.columnNames,
        sheetName,
      },
      row,
      formulaSheet ?? '',
    ),
  )

  const worksheetXml = await pkg.readText(worksheetPath)
  const partPattern = new RegExp(
    `<tablePart\\b[^>]*\\br:id="${escapeRegExp(table.relId)}"[^>]*\\/>`,
  )
  if (!partPattern.test(worksheetXml)) {
    throw new TableEditError(`Table "${table.name}" is not registered on its worksheet.`)
  }
  const cleaned = retallyTableParts(worksheetXml.replace(partPattern, ''))
  pkg.write(worksheetPath, cleaned)
  touchedEntries.add(worksheetPath)

  const relsPath = relsPathFor(worksheetPath)
  if (await pkg.has(relsPath)) {
    const relsXml = await pkg.readText(relsPath)
    const relPattern = new RegExp(
      `<Relationship\\b[^>]*\\bId="${escapeRegExp(table.relId)}"[^>]*\\/?>(?:<\\/Relationship>)?`,
    )
    const updatedRels = relsXml.replace(relPattern, '')
    if (updatedRels !== relsXml) {
      pkg.write(relsPath, updatedRels)
      touchedEntries.add(relsPath)
    }
  }

  const contentTypes = await pkg.readText('[Content_Types].xml')
  const overridePattern = new RegExp(
    `<Override\\b[^>]*PartName="/${escapeRegExp(table.path)}"[^>]*\\/>`,
  )
  if (overridePattern.test(contentTypes)) {
    pkg.write('[Content_Types].xml', contentTypes.replace(overridePattern, ''))
    touchedEntries.add('[Content_Types].xml')
  }

  pkg.remove(table.path)
  touchedEntries.add(table.path)
}

/// Recomputes `<tableParts count>`; drops the element when its last part is
/// gone. Returns the worksheet XML unchanged when there is no tableParts.
function retallyTableParts(worksheetXml: string): string {
  const wrapper = /<tableParts\b[^>]*>[\s\S]*?<\/tableParts>/.exec(worksheetXml)
  if (!wrapper) return worksheetXml
  const remaining = (wrapper[0].match(/<tablePart\b/g) ?? []).length
  if (remaining === 0) return worksheetXml.replace(wrapper[0], '')
  return worksheetXml.replace(wrapper[0], wrapper[0].replace(/count="\d+"/, `count="${remaining}"`))
}

/// Runs `rewrite` over every worksheet `<f>` body and workbook definedName.
/// `rewrite` receives the formula text, the formula cell's 0-based row, and
/// the worksheet's sheet name (workbook-level rewrites pass -1/null).
async function rewriteFormulasEverywhere(
  pkg: MutablePackage,
  touchedEntries: Set<string>,
  rewrite: (formula: string, row: number, sheetName: string | null) => string,
): Promise<void> {
  const sheetNames = await sheetNamesByPath(pkg)
  for (const path of await pkg.paths()) {
    if (!WORKSHEET_PATTERN.test(path)) continue
    const sheetName = sheetNames.get(path) ?? null
    const xml = await pkg.readText(path)
    if (!xml.includes('<f')) continue
    let touched = false
    const rewritten = xml.replace(
      /(<c\b[^>]*\br="([A-Z]+)([0-9]+)"[^>]*>)([\s\S]*?)(<\/c>)/g,
      (full, open: string, _label: string, rowText: string, body: string, close: string) => {
        if (!body.includes('<f')) return full
        const row = Number(rowText) - 1
        const patched = body.replace(
          /(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g,
          (cell, fOpen: string, formula: string, fClose: string) => {
            const next = rewrite(formula, row, sheetName)
            if (next === formula) return cell
            touched = true
            return `${fOpen}${next}${fClose}`
          },
        )
        return touched ? `${open}${patched}${close}` : full
      },
    )
    if (rewritten !== xml) {
      pkg.write(path, rewritten)
      touchedEntries.add(path)
    }
  }
  if (await pkg.has('xl/workbook.xml')) {
    const workbookXml = await pkg.readText('xl/workbook.xml')
    let namesTouched = false
    const rewritten = workbookXml.replace(
      /(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g,
      (full, open: string, formula: string, close: string) => {
        const next = rewrite(formula, -1, null)
        if (next === formula) return full
        namesTouched = true
        return `${open}${next}${close}`
      },
    )
    if (namesTouched) {
      pkg.write('xl/workbook.xml', rewritten)
      touchedEntries.add('xl/workbook.xml')
    }
  }
}

/// Resolves each worksheet part path to its workbook.xml sheet name, so A1
/// rewrites can qualify references from other sheets. Attribute order varies
/// by producer, so sheet elements and relationships are matched in two steps.
async function sheetNamesByPath(pkg: MutablePackage): Promise<Map<string, string>> {
  const map = new Map<string, string>()
  if (!(await pkg.has('xl/workbook.xml')) || !(await pkg.has('xl/_rels/workbook.xml.rels'))) {
    return map
  }
  const workbookXml = await pkg.readText('xl/workbook.xml')
  const relsXml = await pkg.readText('xl/_rels/workbook.xml.rels')
  const relTargets = new Map<string, string>()
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const id = /\bId="([^"]+)"/.exec(match[0])?.[1]
    const target = /\bTarget="([^"]+)"/.exec(match[0])?.[1]
    if (id && target) relTargets.set(id, target)
  }
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/?>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[0])?.[1]
    const relId = /\br:id="([^"]+)"/.exec(match[0])?.[1]
    const target = relId ? relTargets.get(relId) : undefined
    if (name && target) {
      map.set(`xl/${target.replace(/^\/?xl\//, '').replace(/^\.\//, '')}`, decodeEntities(name))
    }
  }
  return map
}

/// The resized area must not overlap another table in the workbook.
async function assertNoOverlapWithOtherTables(
  pkg: MutablePackage,
  tableName: string,
  selfPath: string,
  area: TableArea,
): Promise<void> {
  for (const path of await tablePartPaths(pkg)) {
    if (path === selfPath) continue
    const xml = await pkg.readText(path)
    const ref = /<table\b[^>]*\bref="([^"]+)"/.exec(xml)?.[1]
    if (ref && areasOverlap(area, parseRef(ref))) {
      throw new TableEditError(
        `Table "${tableName}" would overlap an existing table (${ref}) — shrink it first.`,
      )
    }
  }
}

/// Excel forbids a table overlapping the sheet auto-filter or merged cells.
async function assertNoSheetConflicts(
  pkg: MutablePackage,
  worksheetPath: string,
  tableName: string,
  area: TableArea,
): Promise<void> {
  const worksheetXml = await pkg.readText(worksheetPath)
  const autoFilterRef = /<autoFilter\b[^>]*\bref="([^"]+)"/.exec(worksheetXml)?.[1]
  if (autoFilterRef && areasOverlap(area, parseRef(autoFilterRef))) {
    throw new TableEditError(
      `Table "${tableName}" would overlap the sheet auto-filter (${autoFilterRef}) — clear it first.`,
    )
  }
  for (const match of worksheetXml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)) {
    const ref = match[1]
    if (ref && areasOverlap(area, parseRef(ref))) {
      throw new TableEditError(
        `Table "${tableName}" would overlap merged cells (${ref}) — unmerge them first.`,
      )
    }
  }
}

function areasOverlap(a: TableArea, b: TableArea): boolean {
  return (
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startColumn <= b.endColumn &&
    b.startColumn <= a.endColumn
  )
}

async function assertNameNotDefined(pkg: MutablePackage, name: string): Promise<void> {
  const workbookXml = await pkg.readText('xl/workbook.xml')
  for (const match of workbookXml.matchAll(/<definedName\b[^>]*\bname="([^"]+)"/g)) {
    if (match[1]?.toLowerCase() === name.toLowerCase()) {
      throw new TableEditError(`Table name "${name}" collides with a defined name in the workbook.`)
    }
  }
}

function decodeEntities(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function escapeAttribute(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
