/// Drag & drop a local document file anywhere in the app window to open it.
///
/// Renderers share one main process (the shell hosts every editor view), so the
/// bridge only has to resolve dropped files to paths and send them over ONE
/// channel; the shell routes each path through its normal open pipeline. The
/// installer runs inside preloads, where `webUtils.getPathForFile` is callable
/// and DOM listeners see real OS drops — paths never cross contextBridge.
///
/// Ownership contract with page-level drop zones: a zone that handles its own
/// file drops (AI attachment panels, slides image insert) cancels the event
/// before it bubbles here (`defaultPrevented`), and this bridge stays out of
/// the way. Text/media drags without OS files are untouched.
import { ipcRenderer, webUtils } from 'electron'

import { WITNESS_DROP_CHANNEL } from './witnessed-drops'

export const DROP_OPEN_CHANNEL = 'app:open-dropped-files'

/** Extensions routed by apps/shell routeDocumentPath — keep in sync there and
 *  with OPEN_DIALOG_EXTENSIONS / OPEN_LOCAL_EXTENSIONS on the home screen. */
export const OPENABLE_DOC_RE = /\.(docx|xlsx|xlsm|xls|csv|pptx|pdf|md|markdown|html|htm)$/i

/** Recognized-but-unsupported formats: kept in the sent payload so the shell
 *  can show its "not supported" dialog instead of dropping them silently.
 *  Mirrors UNSUPPORTED_DOC_RE in apps/shell/src/main/index.ts. */
export const KNOWN_UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsb|pages|key|numbers)$/i

/** upper bound on how many files one drop may ask to open */
const MAX_DROPPED_FILES = 20

/** the resolver signature webUtils.getPathForFile satisfies; injectable for tests */
type PathResolver = (file: File) => string

/** Resolve one dropped file, tolerating resolver failures (see droppableFilePaths). */
function tryResolvePath(file: File, getPathForFile: PathResolver): string {
  try {
    return getPathForFile(file).trim()
  } catch {
    return ''
  }
}

/**
 * Resolve an event's dropped files to local paths. Returns null when the drag
 * carries no OS files at all (internal text/element drags), or [] when it does
 * but none resolve (directories, virtual entries) — both mean "not ours".
 */
export function droppableFilePaths(
  ev: Pick<DragEvent, 'dataTransfer'>,
  getPathForFile: PathResolver,
): string[] | null {
  const transfer = ev.dataTransfer
  if (!transfer || !transfer.types.includes('Files')) return null
  const paths: string[] = []
  for (const file of Array.from(transfer.files)) {
    // A throwing resolver (e.g. a sandboxed entry Electron cannot map) must
    // not abort the whole drop: skip that file like a virtual entry.
    // Non-empty guard covers virtual entries (e.g. page-referenced blobs).
    const path = tryResolvePath(file, getPathForFile)
    if (path) paths.push(path)
  }
  return paths
}

/**
 * Sanitize + classify a raw IPC payload: strings only, trimmed, deduped,
 * capped, then split into directly-openable paths and known-unsupported
 * extensions (unique, first-seen order). Anything unrecognized (.png, .zip,
 * nonexistent junk from a hostile sender) falls out silently.
 */
export function partitionDropPayload(raw: unknown): {
  supported: string[]
  unsupportedExts: string[]
} {
  const supported: string[] = []
  const unsupportedExts: string[] = []
  if (!Array.isArray(raw)) return { supported, unsupportedExts }
  const seen = new Set<string>()
  for (const entry of raw as unknown[]) {
    if (typeof entry !== 'string') continue
    const path = entry.trim()
    if (!path || seen.has(path)) continue
    seen.add(path)
    if (OPENABLE_DOC_RE.test(path)) {
      // keep scanning for unsupported entries even after hitting the open cap
      if (supported.length < MAX_DROPPED_FILES) supported.push(path)
    } else if (KNOWN_UNSUPPORTED_DOC_RE.test(path)) {
      const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
      if (!unsupportedExts.includes(ext)) unsupportedExts.push(ext)
    }
  }
  return { supported, unsupportedExts }
}

/** Symbol.for keeps repeated installs idempotent when several bundled copies of
 *  this module end up in one process (mirrors navigation-guard). */
const INSTALLED = Symbol.for('airy.drop-open-installed')

/**
 * Preload-side hook: makes `drop` fire for document drags anywhere in the
 * page and forwards recognized files to the shell's router. Intentionally not
 * exposed through contextBridge — page code cannot spoof these events with
 * arbitrary paths because payload construction happens here in the preload
 * world only, and every drop/paste listener ignores synthetic (page-
 * dispatched) events via the isTrusted check: only real user input carries
 * isTrusted=true, so a page that grabs a File from a DOM input cannot forge
 * a witness or an open by dispatching its own events.
 */
export function installDropOpenBridge(): void {
  const holder = globalThis as Record<symbol, boolean | undefined>
  if (typeof window === 'undefined' || holder[INSTALLED]) return
  holder[INSTALLED] = true

  // Consent witness for the files:add attachment channel: record every OS
  // file that really lands in this window — including drops a page drop zone
  // claims (capture runs before page handlers and only observes). The main
  // process lets files:add widen its read allowlist only for witnessed paths
  // (see witnessed-drops.ts), so a compromised page cannot self-grant; the
  // isTrusted guard below keeps the witness itself unforgeable (synthetic
  // drop/paste events dispatched by page script never become witnesses).
  const witness = (files: FileList | null | undefined): void => {
    if (!files || files.length === 0) return
    const paths: string[] = []
    for (const file of Array.from(files)) {
      const path = tryResolvePath(file, (f) => webUtils.getPathForFile(f))
      if (path) paths.push(path)
    }
    if (paths.length > 0) ipcRenderer.send(WITNESS_DROP_CHANNEL, paths.slice(0, 20))
  }
  // real user events are trusted; synthetic ones (page-dispatched) are not
  window.addEventListener(
    'drop',
    (ev) => {
      if (!ev.isTrusted) return
      witness(ev.dataTransfer?.files)
    },
    { capture: true },
  )
  // Files pasted with a local path (a copied file) ride the same witness.
  window.addEventListener(
    'paste',
    (ev) => {
      if (!ev.isTrusted) return
      witness(ev.clipboardData?.files)
    },
    { capture: true },
  )

  // Without a canceled dragover Chromium never fires `drop`; canceling here is
  // what lets a document-drag land anywhere that isn't already a drop zone.
  window.addEventListener('dragover', (ev) => {
    if (ev.defaultPrevented) return
    if (!ev.dataTransfer?.types.includes('Files')) return
    ev.preventDefault()
  })

  window.addEventListener('drop', (ev) => {
    // first: something in the page already claimed this drop (image insert,
    // AI attachments...) — never second-guess it. Synthetic drops are ignored
    // for the same reason as in the witness above.
    if (ev.defaultPrevented || !ev.isTrusted) return
    const paths = droppableFilePaths(ev, (file) => webUtils.getPathForFile(file))
    if (!paths) return
    // only recognized documents ride along; stray images/folders are swallowed
    // here so they neither navigate the page nor produce open attempts
    const payload = paths.filter((p) => OPENABLE_DOC_RE.test(p) || KNOWN_UNSUPPORTED_DOC_RE.test(p))
    ev.preventDefault()
    if (payload.length === 0) return
    ipcRenderer.send(DROP_OPEN_CHANNEL, payload.slice(0, MAX_DROPPED_FILES))
  })
}
