/// BUG-1603 print guard: a sheet saved at a fixed 100% scale on paper the
/// user does not expect (UR-04: a 14-column table on Letter) tiles into N
/// horizontal strips and reads as "prints however it wants", even though the
/// pipeline faithfully executes the file's page setup. Excel only fits a
/// sheet when asked, so the guard stays a suggestion: it detects the pain
/// and offers one-click fixes that override the current print job without
/// touching the saved file.
import type { Lang } from '@airy-office/i18n'

/// OOXML paper-size codes (print-html's PAPER_SIZES maps them to page sizes).
export const PAPER_LETTER = 1
export const PAPER_A4 = 9

/// The paper the user's locale expects: Letter for the US-oriented English
/// UI, A4 for every other UI language (the same split Excel's locale
/// defaults follow for fresh workbooks).
export function localePaperSize(lang: Lang): number {
  return lang === 'en' ? PAPER_LETTER : PAPER_A4
}

/// What the print pipeline measured for the job as currently configured.
export interface PrintGuardInput {
  /// Fit-to-page is on: the fit scale already constrains the width, so the
  /// strip suggestion never applies.
  readonly fitToPage: boolean
  /// Column stripes the job prints across at its fixed scale (the
  /// print-html stripe model); null when the caller did not measure
  /// (fit-to-page jobs have no fixed scale).
  readonly stripsAcross: number | null
  /// The job's paper (OOXML code).
  readonly paperSize: number
  /// localePaperSize of the UI language.
  readonly localePaperSize: number
}

/// Resolved suggestion the dialog renders; every field is a job-scoped
/// override, never a file edit.
export interface PrintGuard {
  /// Stripes at the fixed scale — the "{n} pages wide" figure of the hint;
  /// null when not measured.
  readonly stripsAcross: number | null
  readonly suggestFitToWidth: boolean
  readonly suggestPaper: boolean
  /// The paper code the paper suggestion switches to.
  readonly localePaperSize: number
}

export function detectPrintPain(input: PrintGuardInput): PrintGuard {
  return {
    stripsAcross: input.stripsAcross,
    suggestFitToWidth: !input.fitToPage && input.stripsAcross !== null && input.stripsAcross > 1,
    suggestPaper: input.paperSize !== input.localePaperSize,
    localePaperSize: input.localePaperSize,
  }
}
