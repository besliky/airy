import { PDFName, PDFNumber, PDFRef } from 'pdf-lib'
import type { PDFDocument, PDFObject, PDFPage } from 'pdf-lib'

const ROTATE_KEY = PDFName.of('Rotate')

/**
 * Snap an arbitrary /Rotate value to the nearest valid one per the PDF spec
 * (a multiple of 90): 45 → 90, 225 → 270, -90 → 270, 400 → 0. Non-finite
 * input reads as 0.
 */
export function snapRotation(angle: number): number {
  if (!Number.isFinite(angle)) return 0
  return (((Math.round(angle / 90) * 90) % 360) + 360) % 360
}

/**
 * Read a page's effective /Rotate without ever throwing. PDFPage.getRotation()
 * crashes with UnexpectedObjectTypeError on files whose producer wrote a
 * non-numeric value (e.g. `/Rotate /90` — a PDFName instead of the integer),
 * which used to turn every annotate+save of such files into an opaque failure.
 * Viewers (Adobe, pdf.js, poppler) tolerate the malformed value and display
 * the page unrotated, so every save-pipeline read goes through here instead:
 * missing or non-numeric values (including null and unresolvable refs) read
 * as 0, numbers snap to the nearest valid multiple of 90.
 */
export function pageRotation(page: PDFPage): number {
  try {
    const raw: PDFObject | undefined = page.node.getInheritableAttribute(ROTATE_KEY)
    // Resolve indirect refs manually: context.lookup(raw) with no type
    // arguments resolves refs (and maps null to undefined) without asserting
    const value = raw instanceof PDFRef ? page.node.context.lookup(raw) : raw
    return value instanceof PDFNumber ? snapRotation(value.asNumber()) : 0
  } catch {
    // Unreadable attribute chain (e.g. malformed /Parent): no rotation
    return 0
  }
}

/**
 * Replace malformed effective /Rotate values (anything non-numeric, e.g. the
 * `/Rotate /90` names written by broken producers) with a numeric 0 on the
 * page itself, so a saved file is spec-valid instead of perpetuating the
 * defect and future reads never see the broken value. Valid rotations
 * (numeric, including indirect refs) are left untouched. Called after
 * loading the document in the save pipeline.
 */
export function repairBrokenPageRotations(pdfDoc: PDFDocument): void {
  for (const page of pdfDoc.getPages()) {
    try {
      const raw = page.node.getInheritableAttribute(ROTATE_KEY)
      const value = raw instanceof PDFRef ? page.node.context.lookup(raw) : raw
      if (value !== undefined && !(value instanceof PDFNumber)) {
        page.node.set(ROTATE_KEY, PDFNumber.of(0))
      }
    } catch {
      // Unreadable attribute chain: force a known-good value on the leaf
      page.node.set(ROTATE_KEY, PDFNumber.of(0))
    }
  }
}
