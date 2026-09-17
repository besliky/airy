/**
 * Registry of docs renderer webContents (tab views and standalone windows).
 *
 * The shell bundles every editor's main code into one process, so the
 * process-global win:new channel answers any webContents. Only docs
 * renderers legitimately call it (grep: the docs ribbon's New Tab button is
 * the sole caller), so the handler validates the sender against this set.
 * Track/untrack happen where docs webContents are created and torn down;
 * the pure membership test keeps the policy unit-testable.
 */
const docsRendererIds = new Set<number>()

export function trackDocsRenderer(wcId: number): void {
  docsRendererIds.add(wcId)
}

export function untrackDocsRenderer(wcId: number): void {
  docsRendererIds.delete(wcId)
}

export function isDocsRenderer(senderId: number | undefined): boolean {
  return senderId !== undefined && docsRendererIds.has(senderId)
}

/** Test helper. */
export function resetDocsRenderers(): void {
  docsRendererIds.clear()
}
