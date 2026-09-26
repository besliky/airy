/// Table-slicer persistence (PAR-203): persists a slicer bound to a table
/// column as the OOXML parts Excel writes — `xl/slicers/slicerN.xml` (the
/// visible control), `xl/slicerCaches/slicerCacheN.xml` (with the x15
/// `tableSlicerCache` extension naming tableId + column), the workbook's x15
/// `slicerCaches` ext with its slicerCache relationship plus the hidden
/// `Slicer_*` defined name, and the worksheet's x15 `slicerList` ext with its
/// slicer relationship. The filter criteria themselves are NOT stored here —
/// they live in the table part's own autoFilter (see xlsx-filter.ts), so a
/// saved file keeps filtering even for readers that ignore slicer parts.
/// The visual drawing anchor (the on-sheet graphicFrame shape) is not
/// written: the app renders slicers as panels, and Excel shows the file's
/// filter state without the slicer button.

import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeOverride,
  relativeTarget,
  relsPathFor,
} from './xlsx-drawing-add'
import { findTablePart } from './xlsx-table-edit'
import type { MutablePackage } from './xlsx-drawing-add'

export class SlicerAddError extends Error {}

export interface TableSlicerAddition {
  readonly worksheetPath: string
  /// The table's displayName token.
  readonly tableName: string
  /// 0-based offset into the table's tableColumns list.
  readonly colId: number
}

const SLICER_CONTENT_TYPE = 'application/vnd.ms-excel.slicer+xml'
const SLICER_CACHE_CONTENT_TYPE = 'application/vnd.ms-excel.slicerCache+xml'
const SLICER_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const SLICER_CACHE_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
/// x15 (2010/11) namespaces — table slicers bind through the x15 extensions,
/// unlike pivot slicers which use x14.
const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
const MS_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const TABLE_SLICER_CACHE_EXT_URI = '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}'
const WORKBOOK_SLICER_CACHES_EXT_URI = '{46BE6895-7355-4a93-B00E-2C351335B9C9}'
const WORKSHEET_SLICER_LIST_EXT_URI = '{3A4CF648-6AED-4f9d-8E56-E3E7D4C4D3F7}'
const DEFAULT_ROW_HEIGHT_EMU = 241300

interface TableBinding {
  /// table/@id — the token x15:tableSlicerCache/@tableId references.
  readonly tableId: number
  readonly columnNames: readonly string[]
}

export async function applySlicerAdditions(
  pkg: MutablePackage,
  additions: readonly TableSlicerAddition[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (additions.length === 0) return
  // Names already taken by file slicers and defined names (the slicer cache
  // name shares Excel's defined-name namespace).
  const takenSlicers = await collectSlicerNames(pkg)
  const takenCaches = new Set([
    ...definedNames(await pkg.readText('xl/workbook.xml')),
    ...takenSlicers,
  ])
  for (const addition of additions) {
    const table = await findTablePart(pkg, addition.worksheetPath, addition.tableName)
    if (!table) {
      throw new SlicerAddError(`Table "${addition.tableName}" was not found on this sheet.`)
    }
    const binding = await tableBinding(pkg, table.path)
    if (addition.colId < 0 || addition.colId >= binding.columnNames.length) {
      throw new SlicerAddError(
        `Slicer column ${addition.colId + 1} is outside table "${addition.tableName}".`,
      )
    }
    // Skip when the package already carries an identical binding (the same
    // table column sliced twice would corrupt the filter criteria ownership).
    const columnName = binding.columnNames[addition.colId]!
    if (await hasTableSlicerBinding(pkg, binding.tableId, addition.colId + 1)) {
      continue
    }
    const cacheName = uniqueName(`Slicer_${sanitizeCacheName(columnName)}`, takenCaches)
    const slicerName = uniqueName(columnName, takenSlicers)
    takenCaches.add(cacheName)
    takenSlicers.add(slicerName)

    const slicerPath = await allocatePartPath(pkg, 'xl/slicers/slicer', '.xml')
    pkg.add(slicerPath, buildSlicerXml(slicerName, cacheName, columnName))
    touchedEntries.add(slicerPath)
    await registerContentTypeOverride(pkg, slicerPath, SLICER_CONTENT_TYPE, touchedEntries)

    const cachePath = await allocatePartPath(pkg, 'xl/slicerCaches/slicerCache', '.xml')
    pkg.add(
      cachePath,
      buildSlicerCacheXml(cacheName, columnName, binding.tableId, addition.colId + 1),
    )
    touchedEntries.add(cachePath)
    await registerContentTypeOverride(pkg, cachePath, SLICER_CACHE_CONTENT_TYPE, touchedEntries)

    const cacheRelId = await appendRelationship(
      pkg,
      relsPathFor('xl/workbook.xml'),
      SLICER_CACHE_REL_TYPE,
      relativeTarget('xl/workbook.xml', cachePath),
    )
    touchedEntries.add(relsPathFor('xl/workbook.xml'))

    const slicerRelId = await appendRelationship(
      pkg,
      relsPathFor(addition.worksheetPath),
      SLICER_REL_TYPE,
      relativeTarget(addition.worksheetPath, slicerPath),
    )
    touchedEntries.add(relsPathFor(addition.worksheetPath))

    pkg.write('xl/workbook.xml', await attachWorkbookSlicerCache(pkg, cacheName, cacheRelId))
    touchedEntries.add('xl/workbook.xml')
    pkg.write(
      addition.worksheetPath,
      attachWorksheetSlicerList(await pkg.readText(addition.worksheetPath), slicerRelId),
    )
    touchedEntries.add(addition.worksheetPath)
  }
}

async function tableBinding(pkg: MutablePackage, tablePath: string): Promise<TableBinding> {
  const open = /<table\b[^>]*>/.exec(await pkg.readText(tablePath))?.[0] ?? ''
  const id = Number(/\bid="(\d+)"/.exec(open)?.[1])
  const columnNames = [...(await pkg.readText(tablePath)).matchAll(/<tableColumn\b[^>]*\/>/g)].map(
    (column) => decodeEntities(/\bname="([^"]*)"/.exec(column[0])?.[1] ?? ''),
  )
  if (!Number.isInteger(id) || columnNames.length === 0) {
    throw new SlicerAddError('The table part has no usable id or column list.')
  }
  return { tableId: id, columnNames }
}

/// True when the workbook already carries a slicerCache whose
/// x15 tableSlicerCache names the same tableId + 1-based column.
async function hasTableSlicerBinding(
  pkg: MutablePackage,
  tableId: number,
  column: number,
): Promise<boolean> {
  for (const path of await slicerCachePaths(pkg)) {
    const xml = await pkg.readText(path)
    if (!xml.includes(`tableId="${tableId}"`)) continue
    const pattern = new RegExp(
      `<x15:tableSlicerCache[^>]*\\btableId="${tableId}"[^>]*\\bcolumn="${column}"`,
    )
    if (pattern.test(xml)) return true
  }
  return false
}

async function slicerCachePaths(pkg: MutablePackage): Promise<string[]> {
  return (await pkg.paths()).filter((path) => /^xl\/slicerCaches\/[^/]+\.xml$/.test(path)).sort()
}

async function collectSlicerNames(pkg: MutablePackage): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of (await pkg.paths()).filter((path) => /^xl\/slicers\/[^/]+\.xml$/.test(path))) {
    for (const match of (await pkg.readText(path)).matchAll(/<slicer\b[^>]*\bname="([^"]+)"/g)) {
      if (match[1]) names.add(match[1])
    }
  }
  return names
}

function definedNames(workbookXml: string): string[] {
  return [...workbookXml.matchAll(/<definedName\b[^>]*\bname="([^"]+)"/g)]
    .map((match) => match[1])
    .filter((name): name is string => name !== undefined)
}

/// Excel's slicer-cache name shape: letters kept, digits/dots kept after the
/// first character, everything else becomes an underscore.
function sanitizeCacheName(name: string): string {
  let sanitized = ''
  for (const [index, char] of [...name].entries()) {
    sanitized += /[A-Za-z]/.test(char) || (index > 0 && /[0-9.]/.test(char)) ? char : '_'
  }
  return sanitized || 'Column'
}

function uniqueName(candidate: string, taken: ReadonlySet<string>): string {
  if (!taken.has(candidate)) return candidate
  let suffix = 1
  while (taken.has(`${candidate} ${suffix}`)) suffix += 1
  return `${candidate} ${suffix}`
}

function buildSlicerXml(name: string, cacheName: string, caption: string): string {
  return (
    xmlDeclaration() +
    `<slicers xmlns="${MS_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="x" xmlns:x="${MAIN_NS}">` +
    `<slicer name="${escapeAttribute(name)}" cache="${escapeAttribute(cacheName)}" ` +
    `caption="${escapeAttribute(caption)}" rowHeight="${DEFAULT_ROW_HEIGHT_EMU}"/>` +
    '</slicers>'
  )
}

function buildSlicerCacheXml(
  name: string,
  sourceName: string,
  tableId: number,
  column: number,
): string {
  return (
    xmlDeclaration() +
    `<slicerCacheDefinition xmlns="${MS_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="x" ` +
    `xmlns:x="${MAIN_NS}" name="${escapeAttribute(name)}" ` +
    `sourceName="${escapeAttribute(sourceName)}">` +
    '<extLst>' +
    `<ext uri="${TABLE_SLICER_CACHE_EXT_URI}" xmlns:x15="${X15_NS}">` +
    '<x15:slicerCacheDefinition>' +
    `<x15:tableSlicerCache tableId="${tableId}" column="${column}"/>` +
    '</x15:slicerCacheDefinition>' +
    '</ext>' +
    '</extLst>' +
    '</slicerCacheDefinition>'
  )
}

/// Appends the x15 slicerCaches ext to the workbook's extLst and records the
/// `Slicer_*` defined name Excel pairs with every slicer cache.
async function attachWorkbookSlicerCache(
  pkg: MutablePackage,
  cacheName: string,
  cacheRelId: string,
): Promise<string> {
  let workbookXml = await pkg.readText('xl/workbook.xml')
  const cachesElement =
    `<ext uri="${WORKBOOK_SLICER_CACHES_EXT_URI}" xmlns:x15="${X15_NS}">` +
    `<x15:slicerCaches><x15:slicerCache r:id="${cacheRelId}"/></x15:slicerCaches>` +
    '</ext>'
  // The slicerCache defined name refers to #N/A (Excel's marker for
  // "cache-backed, not address-backed"); workbook-level scope.
  const definedNameElement = `<definedName name="${escapeAttribute(cacheName)}">#N/A</definedName>`
  if (workbookXml.includes('</definedNames>')) {
    workbookXml = workbookXml.replace(
      '</definedNames>',
      () => `${definedNameElement}</definedNames>`,
    )
  } else {
    // No definedNames section yet: schema order puts one after sheets/before
    // the calcPr-ish trailing elements — insert before </workbook> is the
    // tolerant spot Excel accepts.
    const closeAt = workbookXml.lastIndexOf('</workbook>')
    if (closeAt < 0) throw new SlicerAddError('workbook.xml is malformed.')
    workbookXml =
      workbookXml.slice(0, closeAt) +
      `<definedNames>${definedNameElement}</definedNames>` +
      workbookXml.slice(closeAt)
  }
  const workbookExtAt = workbookXml.lastIndexOf('</extLst>')
  if (workbookExtAt >= 0) {
    return workbookXml.slice(0, workbookExtAt) + cachesElement + workbookXml.slice(workbookExtAt)
  }
  const closeAt = workbookXml.lastIndexOf('</workbook>')
  if (closeAt < 0) throw new SlicerAddError('workbook.xml is malformed.')
  return (
    workbookXml.slice(0, closeAt) + `<extLst>${cachesElement}</extLst>` + workbookXml.slice(closeAt)
  )
}

/// Appends the x15 slicerList ext to the worksheet's extLst (creating the
/// worksheet-level extLst when absent — only the one past sheetData counts).
function attachWorksheetSlicerList(worksheetXml: string, slicerRelId: string): string {
  const listElement =
    `<ext uri="${WORKSHEET_SLICER_LIST_EXT_URI}" xmlns:x15="${X15_NS}">` +
    `<x15:slicerList><x15:slicer r:id="${slicerRelId}"/></x15:slicerList>` +
    '</ext>'
  const worksheetExtAt = worksheetXml.lastIndexOf('</extLst>')
  if (worksheetExtAt >= 0 && worksheetExtAt > worksheetXml.lastIndexOf('</sheetData>')) {
    return worksheetXml.slice(0, worksheetExtAt) + listElement + worksheetXml.slice(worksheetExtAt)
  }
  const closeAt = worksheetXml.lastIndexOf('</worksheet>')
  if (closeAt < 0) throw new SlicerAddError('The worksheet part is malformed.')
  return (
    worksheetXml.slice(0, closeAt) + `<extLst>${listElement}</extLst>` + worksheetXml.slice(closeAt)
  )
}

function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
}

function decodeEntities(input: string): string {
  return input
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
}

function escapeAttribute(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
