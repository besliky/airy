/// Page bookkeeping for Excel's differentFirst / differentOddEven header
/// and footer variants. Chromium's printToPDF takes ONE header/footer
/// template pair per call, so the export prints the whole sheet with the
/// odd-page templates and then, when a variant is active, re-prints page 1
/// (first-page templates) and/or the even pages (even templates) with
/// `pageRanges`; the final PDF takes each page from the pass that owns it.
/// Chromium numbers the pages of a ranged print by their document index, so
/// `&P`/`&N` stay correct in every pass. Pure — no Electron here.

export type PageVariant = 'odd' | 'even' | 'first'

export interface VariantFlags {
  readonly hasFirst: boolean
  readonly hasEven: boolean
}

/// Which pass prints the given 1-based page.
export function variantForPage(pageNumber: number, flags: VariantFlags): PageVariant {
  if (pageNumber === 1 && flags.hasFirst) return 'first'
  if (pageNumber % 2 === 0 && flags.hasEven) return 'even'
  return 'odd'
}

/// `pageRanges` for the even pass of a `total`-page document ('' when none).
export function evenPageRanges(total: number): string {
  const pages: string[] = []
  for (let page = 2; page <= total; page += 2) pages.push(String(page))
  return pages.join(',')
}

export interface StitchStep {
  /// 1-based document page.
  readonly page: number
  readonly source: PageVariant
  /// Page index within the source pass's PDF.
  readonly index: number
}

/// Page-by-page assembly plan: the odd pass holds every page at its
/// document index; the first pass is page 1 alone; the even pass holds
/// pages 2, 4, 6, … in order.
export function stitchPlan(total: number, flags: VariantFlags): StitchStep[] {
  const steps: StitchStep[] = []
  for (let page = 1; page <= total; page += 1) {
    const source = variantForPage(page, flags)
    steps.push({
      page,
      source,
      index: source === 'first' ? 0 : source === 'even' ? page / 2 - 1 : page - 1,
    })
  }
  return steps
}

/// One ranged printToPDF pass of a per-sheet plan: the pages sharing a
/// (sheet, variant) template pair, in document order.
export interface SheetPass {
  readonly sheet: number
  readonly variant: PageVariant
  readonly pages: readonly number[]
}

/// Where a document page comes from: the owning pass and its index within
/// that pass (null → the base pass, which holds every page at its index).
export interface SheetPlanStep {
  readonly pass: number
  readonly index: number
}

export interface SheetPlan {
  readonly passes: readonly SheetPass[]
  readonly steps: readonly (SheetPlanStep | null)[]
}

/// Per-sheet pass planning for entire-workbook jobs (BUG-1105): Chromium
/// prints one header/footer template pair per call, but `&A` must name the
/// sheet owning each page. Every page belongs to one sheet (by the
/// renderer's per-sheet page counts — the last sheet also owns any pages
/// past the counted total, so layout-count drift degrades to a neighbouring
/// name instead of losing pages) and to one page variant; pages sharing a
/// (sheet, variant) pair print together in one ranged pass, and every pass
/// numbers its pages by document index, so `&P`/`&N` stay correct. Pure —
/// no Electron here.
export function sheetPassesPlan(
  total: number,
  sheetPages: readonly number[],
  flags: VariantFlags,
): SheetPlan {
  const ownerOf = (page: number): number => {
    let start = 1
    for (let sheet = 0; sheet < sheetPages.length; sheet += 1) {
      const count = sheetPages[sheet] ?? 0
      // a zero-count sheet owns nothing; the last sheet absorbs the tail
      if (page < start + count) return sheet
      if (sheet === sheetPages.length - 1) return sheet
      start += count
    }
    return sheetPages.length - 1
  }
  const passes: Array<{ sheet: number; variant: PageVariant; pages: number[] }> = []
  const stepByPage = new Map<number, SheetPlanStep>()
  for (let page = 1; page <= total; page += 1) {
    const sheet = ownerOf(page)
    const variant = variantForPage(page, flags)
    let passIndex = passes.findIndex((p) => p.sheet === sheet && p.variant === variant)
    if (passIndex === -1) {
      passes.push({ sheet, variant, pages: [] })
      passIndex = passes.length - 1
    }
    const pass = passes[passIndex]!
    pass.pages.push(page)
    stepByPage.set(page, { pass: passIndex, index: pass.pages.length - 1 })
  }
  const steps: (SheetPlanStep | null)[] = []
  for (let page = 1; page <= total; page += 1) steps.push(stepByPage.get(page) ?? null)
  return { passes, steps }
}

/// Ascending pages → a Chromium pageRanges string, consecutive runs folded
/// into "first-last" ('1-3,7,9-12'); '' when there is nothing to print.
export function pageRangesString(pages: readonly number[]): string {
  if (pages.length === 0) return ''
  const parts: string[] = []
  let start = pages[0]!
  let previous = start
  for (const page of pages.slice(1)) {
    if (page === previous + 1) {
      previous = page
      continue
    }
    parts.push(start === previous ? `${start}` : `${start}-${previous}`)
    start = page
    previous = page
  }
  parts.push(start === previous ? `${start}` : `${start}-${previous}`)
  return parts.join(',')
}
