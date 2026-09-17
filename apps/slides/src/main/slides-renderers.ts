/**
 * Registry of slides renderer webContents (tab views and standalone windows).
 *
 * The shell bundles every editor's main code into one process, so handlers
 * registered here answer any webContents. Two of them hand out per-renderer
 * privileges and must answer slides renderers only: the display-media handler
 * (a live screen stream) and slides:recent (which re-grants the recents'
 * folders to the asking renderer). Track/untrack happen where slides
 * webContents are created and torn down; the pure membership test keeps the
 * policy unit-testable.
 */
const slidesRendererIds = new Set<number>()

export function trackSlidesRenderer(wcId: number): void {
  slidesRendererIds.add(wcId)
}

export function untrackSlidesRenderer(wcId: number): void {
  slidesRendererIds.delete(wcId)
}

export function isSlidesRenderer(senderId: number | undefined): boolean {
  return senderId !== undefined && slidesRendererIds.has(senderId)
}

/** Test helper. */
export function resetSlidesRenderers(): void {
  slidesRendererIds.clear()
}
