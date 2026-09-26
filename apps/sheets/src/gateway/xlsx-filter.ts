/// Writes a declarative filter snapshot into worksheet XML: the
/// `<autoFilter>` element (with per-column value / custom / color criteria)
/// plus row visibility inside the filter's row span. Unsupported criteria
/// fail closed. Color criteria intern a dxf in the stylesheet — OOXML stores
/// only the `dxfId` reference inside `<colorFilter>`.

import type { DxfSink } from './xlsx-cf'

export class FilterEditError extends Error {}

interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface FilterColumnState {
  readonly colId: number
  readonly values?: readonly string[] | undefined
  readonly blank?: boolean | undefined
  readonly customs?:
    | {
        readonly and?: boolean | undefined
        readonly filters: readonly {
          readonly val: string | number
          readonly operator?: string | undefined
        }[]
      }
    | undefined
  /// Resolved color criterion: the renderer translates the model's
  /// colorFilters (dxf color already resolved) into kind + #RRGGBB; the
  /// unrepresentable shapes never reach the wire.
  readonly colorFilter?: { readonly kind: 'fill' | 'font'; readonly color: string } | undefined
}

export interface SheetFilterState {
  readonly sheetName: string
  readonly filter: {
    readonly range: CellArea
    readonly columns: readonly FilterColumnState[]
  } | null
  readonly hiddenRows: readonly number[]
  readonly visibilityRange: CellArea
  /// Present when the filter belongs to a table (the table part's own
  /// autoFilter) rather than the worksheet: the criteria are written into
  /// that table part and no worksheet autoFilter is created.
  readonly tableName?: string | undefined
}

const CUSTOM_OPERATORS = new Set([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])

export function applyFilterState(
  worksheetXml: string,
  state: SheetFilterState,
  dxfs?: DxfSink | undefined,
): string {
  const element = state.filter === null ? '' : serializeAutoFilter(state.filter, dxfs)
  const existing = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.exec(
    worksheetXml,
  )
  let xml = worksheetXml
  if (existing) {
    xml =
      worksheetXml.slice(0, existing.index) +
      element +
      worksheetXml.slice(existing.index + existing[0].length)
  } else if (element !== '') {
    xml = insertAfterSheetData(worksheetXml, element)
  }
  return applyFilterRows(xml, state)
}

/// Writes a TABLE-owned filter snapshot into the table part's XML: criteria
/// go inside the table's own `<autoFilter>` element (its ref is kept); a
/// null filter removes the element entirely, matching what Excel writes for
/// a table with the filter toggled off.
export function applyTableFilterState(
  tableXml: string,
  state: SheetFilterState,
  dxfs?: DxfSink | undefined,
): string {
  const existing = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/.exec(tableXml)
  if (state.filter === null) {
    if (!existing) return tableXml
    return tableXml.slice(0, existing.index) + tableXml.slice(existing.index + existing[0].length)
  }
  const ref = existing ? /\bref="([^"]+)"/.exec(existing[0])?.[1] : undefined
  // Keep the existing element's ref (it tracks the table's area through
  // resizes); a fresh element takes the filter range, which for a
  // table-owned filter is the table's own area.
  const element = serializeAutoFilter(state.filter, dxfs)
  const withRef = ref === undefined ? element : element.replace(/\bref="[^"]+"/, `ref="${ref}"`)
  if (existing) {
    return (
      tableXml.slice(0, existing.index) +
      withRef +
      tableXml.slice(existing.index + existing[0].length)
    )
  }
  // No autoFilter in the part yet: Excel places it right after the <table>
  // open tag (before tableColumns).
  const openEnd = /<table\b[^>]*>/.exec(tableXml)
  if (!openEnd) throw new FilterEditError('The table part is malformed.')
  return (
    tableXml.slice(0, openEnd.index + openEnd[0].length) +
    withRef +
    tableXml.slice(openEnd.index + openEnd[0].length)
  )
}

/// Declarative row visibility for a filter snapshot (worksheet XML side).
export function applyFilterRows(worksheetXml: string, state: SheetFilterState): string {
  return applyRowVisibility(worksheetXml, state.visibilityRange, new Set(state.hiddenRows))
}

function serializeAutoFilter(
  filter: NonNullable<SheetFilterState['filter']>,
  dxfs: DxfSink | undefined,
): string {
  const ref = toRef(filter.range)
  const columns = [...filter.columns]
    .sort((left, right) => left.colId - right.colId)
    .map((column) => serializeFilterColumn(column, dxfs))
    .join('')
  return columns === ''
    ? `<autoFilter ref="${ref}"/>`
    : `<autoFilter ref="${ref}">${columns}</autoFilter>`
}

function serializeFilterColumn(column: FilterColumnState, dxfs: DxfSink | undefined): string {
  const parts: string[] = []
  if (column.values !== undefined || column.blank) {
    const blank = column.blank ? ' blank="1"' : ''
    const values = (column.values ?? [])
      .map((value) => `<filter val="${escapeXmlAttribute(value)}"/>`)
      .join('')
    parts.push(`<filters${blank}>${values}</filters>`)
  }
  if (column.customs) {
    for (const custom of column.customs.filters) {
      if (custom.operator !== undefined && !CUSTOM_OPERATORS.has(custom.operator)) {
        throw new FilterEditError(
          `Filter condition "${custom.operator}" cannot be saved as XLSX yet.`,
        )
      }
    }
    const and = column.customs.and ? ' and="1"' : ''
    const filters = column.customs.filters
      .map((custom) => {
        const operator =
          custom.operator === undefined || custom.operator === 'equal'
            ? ''
            : ` operator="${custom.operator}"`
        return `<customFilter${operator} val="${escapeXmlAttribute(String(custom.val))}"/>`
      })
      .join('')
    parts.push(`<customFilters${and}>${filters}</customFilters>`)
  }
  if (column.colorFilter) {
    if (!dxfs) {
      throw new FilterEditError('A color filter needs the workbook stylesheet to save.')
    }
    // OOXML carries only the dxf reference; the criterion color lives in the
    // interned dxf (fill → patternFill/bgColor, font → font/color — the two
    // shapes Excel writes, and the ones the import side resolves back).
    const dxfXml =
      column.colorFilter.kind === 'fill'
        ? `<dxf><fill><patternFill><bgColor rgb="${requireColor(column.colorFilter.color)}"/></patternFill></fill></dxf>`
        : `<dxf><font><color rgb="${requireColor(column.colorFilter.color)}"/></font></dxf>`
    parts.push(`<colorFilter dxfId="${dxfs.internDxf(dxfXml)}"/>`)
  }
  if (parts.length === 0) return ''
  return `<filterColumn colId="${column.colId}">${parts.join('')}</filterColumn>`
}

/// #RRGGBB → FFRRGGBB (fully opaque ARGB, the dxf rgb encoding).
function requireColor(color: string): string {
  const hex = /^#([0-9a-fA-F]{6})$/.exec(color.trim())
  const channels = hex?.[1]
  if (channels === undefined) {
    throw new FilterEditError(`The filter color "${color}" cannot be saved as XLSX.`)
  }
  return `FF${channels.toUpperCase()}`
}

/// Schema order places autoFilter after sheetData (and after the protection
/// block when present).
function insertAfterSheetData(worksheetXml: string, element: string): string {
  let insertAt = -1
  for (const pattern of [
    /<sheetData\s*\/>|<\/sheetData>/,
    /<sheetProtection\b[^>]*\/>/,
    /<protectedRanges\b[^>]*\/>|<\/protectedRanges>/,
    /<\/scenarios>/,
  ]) {
    const match = pattern.exec(worksheetXml)
    if (match) insertAt = Math.max(insertAt, match.index + match[0].length)
  }
  if (insertAt < 0) throw new FilterEditError('Worksheet has no sheetData element.')
  return worksheetXml.slice(0, insertAt) + element + worksheetXml.slice(insertAt)
}

/// Declarative row visibility inside the filter's data rows: listed rows are
/// hidden, every other row in the span is unhidden (matching how Excel
/// re-evaluates a filter). The header row is never touched.
function applyRowVisibility(
  worksheetXml: string,
  range: CellArea,
  hiddenRows: ReadonlySet<number>,
): string {
  const firstDataRow = range.startRow + 1
  const seen = new Set<number>()
  let xml = worksheetXml.replace(
    /<row\b([^>]*?)(\/>|>)/g,
    (full, attributes: string, close: string) => {
      const rowNumber = /(?:^|\s)r="([0-9]+)"/.exec(attributes)?.[1]
      if (rowNumber === undefined) return full
      const rowIndex = Number(rowNumber) - 1
      if (rowIndex < firstDataRow || rowIndex > range.endRow) return full
      seen.add(rowIndex)
      const withoutHidden = attributes.replace(/\s*hidden="[^"]*"/, '')
      const hidden = hiddenRows.has(rowIndex) ? ' hidden="1"' : ''
      return `<row${withoutHidden}${hidden}${close}`
    },
  )
  const missing = [...hiddenRows]
    .filter(
      (rowIndex) => !seen.has(rowIndex) && rowIndex >= firstDataRow && rowIndex <= range.endRow,
    )
    .sort((left, right) => left - right)
  for (const rowIndex of missing) {
    xml = insertEmptyHiddenRow(xml, rowIndex + 1)
  }
  return xml
}

function insertEmptyHiddenRow(worksheetXml: string, rowNumber: number): string {
  const newRow = `<row r="${rowNumber}" hidden="1"/>`
  const rowStartPattern = /<row\b[^>]*?\br="([1-9][0-9]*)"/g
  let match: RegExpExecArray | null
  while ((match = rowStartPattern.exec(worksheetXml)) !== null) {
    if (Number(match[1]) > rowNumber) {
      return worksheetXml.slice(0, match.index) + newRow + worksheetXml.slice(match.index)
    }
  }
  if (worksheetXml.includes('</sheetData>')) {
    return worksheetXml.replace('</sheetData>', () => `${newRow}</sheetData>`)
  }
  const emptySheetData = /<sheetData\s*\/>/
  if (emptySheetData.test(worksheetXml)) {
    return worksheetXml.replace(emptySheetData, () => `<sheetData>${newRow}</sheetData>`)
  }
  throw new FilterEditError('Worksheet has no sheetData element.')
}

function toRef(range: CellArea): string {
  return (
    `${columnToLetters(range.startColumn)}${range.startRow + 1}` +
    `:${columnToLetters(range.endColumn)}${range.endRow + 1}`
  )
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}
