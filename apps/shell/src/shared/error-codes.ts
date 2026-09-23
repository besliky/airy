/**
 * Map common Node filesystem error codes to a stable friendly-error key both
 * the renderer toasts and the main-process error dialog localize. The key is
 * resolved from `err.code` when present and from the leading errno token in
 * the message otherwise (raw strings from IPC results carry it there).
 */
export type FriendlyErrorKey = 'enoent' | 'eperm' | 'ebusy' | 'emfile' | 'eisdir'

const CODE_TO_KEY: Record<string, FriendlyErrorKey> = {
  ENOENT: 'enoent',
  EACCES: 'eperm',
  EPERM: 'eperm',
  EBUSY: 'ebusy',
  ETXTBSY: 'ebusy',
  EMFILE: 'emfile',
  EISDIR: 'eisdir',
}

/** leading errno token as Node prints it: "ENOENT: no such file or directory, …" */
const ERRNO_IN_MESSAGE = /\b(ENOENT|EACCES|EPERM|EBUSY|ETXTBSY|EMFILE|EISDIR):/

export function friendlyErrorKey(error: unknown): FriendlyErrorKey | null {
  if (!error) return null
  if (typeof error === 'object') {
    const code = (error as { code?: unknown }).code
    if (typeof code === 'string') {
      const key = CODE_TO_KEY[code.toUpperCase()]
      if (key) return key
    }
    const message = (error as { message?: unknown }).message
    if (typeof message === 'string') return fromMessage(message)
    return null
  }
  if (typeof error === 'string') return fromMessage(error)
  return null
}

function fromMessage(message: string): FriendlyErrorKey | null {
  const match = ERRNO_IN_MESSAGE.exec(message)
  return match ? (CODE_TO_KEY[match[1]] ?? null) : null
}
