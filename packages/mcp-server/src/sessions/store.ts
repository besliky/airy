// Unified session store: open_document hands out one handle space for every
// session kind (docx, xlsx, read-only text), and the read/save/close tools
// dispatch on the session's kind. Kept as a plain Map like the original docx
// store — handles are random UUIDs and sessions live for the process
// lifetime or until close_document.
import type { DocxSession } from '../docx/session.js'
import type { TextSession } from './text.js'
import type { XlsxSession } from '../xlsx/session.js'

export type DocumentSession = DocxSession | XlsxSession | TextSession

const sessions = new Map<string, DocumentSession>()

export function storeSession<T extends DocumentSession>(session: T): T {
  sessions.set(session.handle, session)
  return session
}

export function getSession(handle: string): DocumentSession {
  const session = sessions.get(handle)
  if (!session)
    throw new Error(
      `Unknown document handle "${handle}". Open the document first with open_document.`,
    )
  return session
}

export function removeSession(handle: string): boolean {
  return sessions.delete(handle)
}

export function sessionCount(): number {
  return sessions.size
}
