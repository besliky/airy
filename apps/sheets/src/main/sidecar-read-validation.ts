import type { z } from 'zod'

/**
 * Sidecar read results are validated in main before they cross IPC, and a
 * failed validation must never hand zod's JSON issue dump to the renderer:
 * the lazy loader streams raw error.message straight into the status bar, so
 * a wire-shape drift would drench it in unrecognized_keys JSON (BUG-1711).
 * Follow the bridge's plain-error convention instead — one readable sentence
 * per channel (matching the preload's response-validation family), while the
 * zod issues go to the main-process log for diagnosis.
 */
export function parseSidecarReadResult<S extends z.ZodType>(
  schema: S,
  result: unknown,
  failure: string,
): z.output<S> {
  const parsed = schema.safeParse(result)
  if (!parsed.success) {
    console.warn('[sheets] sidecar read failed schema validation:', parsed.error.issues)
    throw new Error(failure)
  }
  return parsed.data
}
