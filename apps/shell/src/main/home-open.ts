/**
 * The Home-screen "open this path" decision (BUG-1677): a click on an
 * unreadable file, a vanished one or a directory used to end in total
 * silence — routing returned false, the handler dropped the result, and no
 * tab, dialog or toast ever appeared. Dependency-injected like
 * open-failure.ts so the wiring stays unit-testable without Electron.
 */
import { reportOpenFailure, type OpenFailureDeps } from './open-failure'

export interface HomeOpenDeps {
  /** route the path to a document tab; false = nothing opened */
  openDocument: (path: string) => boolean
  /** the #154 dialog channel plus its localized title (openFailureDeps) */
  openFailure: OpenFailureDeps
}

/**
 * Open a path clicked in Home; when it produces no tab, explain why through
 * the same error dialog the launch paths use (EACCES/EPERM/ENOENT/EISDIR).
 * A valid file opens exactly as before and no dialog is raised. Returns
 * true only when a tab was opened; the IPC handler may ignore the result.
 */
export function openFromHome(raw: unknown, deps: HomeOpenDeps): boolean {
  if (typeof raw !== 'string') return false
  if (deps.openDocument(raw)) return true
  // an intended open that produced no tab must explain itself (BUG-1677)
  reportOpenFailure(raw, deps.openFailure)
  return false
}
