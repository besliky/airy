/**
 * Chart model helpers for the editor: the insert-modal kinds, spec ↔ display
 * conversion (insert / edit), and the save-time application of chart data
 * edits to the package (chart part caches + the embedded xlsx workbook, so
 * Word's "Edit Data" sheet shows the edited numbers).
 */
import {
  findChartWorkbookPath,
  parseChartPartXml,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  readDocxPartBase64,
  type ChartDisplay,
  type ChartPatch,
  type ChartSeries,
  type NewChart,
} from '@airy-office/docx-engine'

/** chart kinds the Insert Chart dialog can create */
export type InsertableChartKind = NewChart['kind']

export const INSERTABLE_CHART_KINDS: InsertableChartKind[] = [
  'bar',
  'line',
  'pie',
  'area',
  'scatter',
  'bubble',
  'doughnut',
]

/** display kinds the edit dialog can prefill (chartex/other degrade to bar) */
const EDITABLE_DISPLAY_KINDS = new Set(['bar', 'line', 'pie', 'area', 'scatter', 'bubble'])

/** uniform bubble size written when a spec carries no explicit sizes */
const DEFAULT_BUBBLE_SIZE = 100

/** display model for a freshly inserted (not yet saved) chart */
export function displayFromSpec(spec: NewChart): ChartDisplay {
  const scatterLike = spec.kind === 'scatter' || spec.kind === 'bubble'
  const series: ChartSeries[] = spec.series.map((ser) => ({
    name: ser.name,
    values: ser.values,
    ...(scatterLike
      ? {
          xValues: ser.xValues ?? spec.categories.map(categoryNumber),
          ...(spec.kind === 'bubble'
            ? { sizes: ser.sizes ?? ser.values.map(() => DEFAULT_BUBBLE_SIZE) }
            : {}),
        }
      : {}),
  }))
  return {
    partPath: '',
    // doughnut displays as the pie renderer + a hole (matches what the parser
    // reads back from the saved c:doughnutChart part)
    kind: spec.kind === 'doughnut' ? 'pie' : spec.kind,
    ...(spec.kind === 'doughnut' ? { holePct: 50 } : {}),
    ...(spec.title ? { title: spec.title } : {}),
    categories: spec.categories,
    series,
  }
}

/** modal prefill state recovered from an existing chart node */
export interface ChartEditSource {
  kind: InsertableChartKind
  title: string
  categories: string[]
  series: Array<{ name: string; values: (number | null)[]; sizes?: (number | null)[] }>
}

/**
 * Rebuild the modal's edit state from the node's models: genChart (inserted
 * this session, full kind set) wins over the parsed display kind; a pie with
 * a hole reads back as doughnut. `other`-kind charts degrade to bar — their
 * cache edits still apply, the kind picker just cannot change the part.
 */
export function chartEditSource(
  display: ChartDisplay | null,
  gen: NewChart | null,
): ChartEditSource | null {
  if (!display && !gen) return null
  const d = display
  const kind: InsertableChartKind = gen
    ? gen.kind
    : d && EDITABLE_DISPLAY_KINDS.has(d.kind)
      ? d.kind === 'pie' && (d.holePct ?? 0) > 0
        ? 'doughnut'
        : (d.kind as InsertableChartKind)
      : 'bar'
  return {
    kind,
    title: gen?.title ?? d?.title ?? '',
    categories: gen ? [...gen.categories] : [...(d?.categories ?? [])],
    series: (gen?.series ?? d?.series ?? []).map((s) => ({
      name: s.name ?? '',
      values: [...s.values],
      ...(s.sizes ? { sizes: [...s.sizes] } : {}),
    })),
  }
}

/**
 * Apply chart data edits to the package parts: patches each chart part's
 * caches (patchChartPartXml) and rewrites the embedded workbook's Sheet1
 * (patchChartWorkbookXlsxBase64) so cache and "Edit Data" numbers agree.
 */
export async function applyChartEdits(
  originalBytes: Uint8Array,
  chartParts: Record<string, string>,
  patches: Array<{ partPath: string; patch: ChartPatch }>,
): Promise<{ partXml: Record<string, string>; partBinary: Record<string, string> }> {
  const partXml: Record<string, string> = {}
  const partBinary: Record<string, string> = {}
  for (const { partPath, patch } of patches) {
    const originalPart = chartParts[partPath]
    if (!originalPart) continue
    const patchedXml = patchChartPartXml(originalPart, patch)
    partXml[partPath] = patchedXml
    const wbPath = await findChartWorkbookPath(originalBytes, partPath)
    if (!wbPath) continue
    const existingBase64 = await readDocxPartBase64(originalBytes, wbPath)
    if (!existingBase64) continue
    const display = parseChartPartXml(patchedXml, partPath)
    if (!display) continue
    const namedSeries = display.series.map((s, i) => ({
      name: s.name ?? `Series${i + 1}`,
      values: s.values as (number | null)[],
    }))
    const updated = await patchChartWorkbookXlsxBase64(
      existingBase64,
      display.categories,
      namedSeries,
    )
    if (updated) partBinary[wbPath] = updated
  }
  return { partXml, partBinary }
}

/** numeric category text → number; null for labels (used as scatter x values) */
function categoryNumber(v: string): number | null {
  const t = v.trim()
  if (t === '') return null
  const n = Number(t)
  return Number.isFinite(n) ? n : null
}
