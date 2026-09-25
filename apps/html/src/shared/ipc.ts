import type { AiPanelPrefs } from '@airy-office/ui'
import type { Lang } from '@airy-office/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@airy-office/ai-provider'

export const HTML_CHANNELS = {
  consumePending: 'html:consume-pending',
  previewUpdate: 'html:preview-update',
  previewInfo: 'html:preview-info',
  presentFullScreen: 'html:present-fullscreen',
  presentNewTab: 'html:present-new-tab',
  readFile: 'html:read-file',
  setEncoding: 'html:set-encoding',
  getEditorPrefs: 'html:get-editor-prefs',
  setEditorPrefs: 'html:set-editor-prefs',
  writeRecovery: 'html:write-recovery',
  save: 'html:save',
  saveRequest: 'html:save-request',
  saveRequestAck: 'html:save-request-ack',
  dirtyChanged: 'html:dirty-changed',
  closeSaveRequest: 'html:close-save-request',
  closeSaveResult: 'html:close-save-result',
  fileRenamed: 'html:file-renamed',
  provisionalTitle: 'html:provisional-title',
  pickImage: 'html:pick-image',
  saveImage: 'html:save-image',
  readImage: 'html:read-image',
  fetchImage: 'html:fetch-image',
  exportRequest: 'html:export-request',
  exportDocx: 'html:export-docx',
  exportPdf: 'html:export-pdf',
  printRequest: 'html:print-request',
  aiGenerateImage: 'html:ai-generate-image',
  filesPick: 'html:files-pick',
  filesAdd: 'html:files-add',
  filesAddPastedImage: 'html:files-add-pasted-image',
  filesRead: 'html:files-read',
  filesReadImage: 'html:files-read-image',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
  getAutoSaveDefault: 'app:get-auto-save-default',
  autoSaveDefaultChanged: 'app:auto-save-default-changed',
  getAiPanelPrefs: 'app:get-ai-panel-prefs',
  aiPanelPrefsChanged: 'app:ai-panel-prefs-changed',
} as const

export type UiTheme = 'light' | 'dark' | 'system'

/**
 * Charsets a manual "Reopen with encoding" pick may name (UX-1696). Mirrors
 * the detector's candidate set (LEGACY_CHARSETS in
 * packages/file-parse/src/text.ts) plus the UTF-8/16 family a user may want
 * to force over a wrong legacy guess. One list shared by the renderer (picker
 * options), the preload (pass-through) and main (persistence + validation).
 * Keep in sync with LEGACY_CHARSETS; a persisted entry whose encoding is not
 * listed here is dropped on read and the file falls back to auto-detection.
 */
export const SELECTABLE_ENCODINGS = [
  'utf-8',
  'utf-16le',
  'utf-16be',
  'gb18030',
  'shift_jis',
  'big5',
  'euc-kr',
  'windows-1252',
  'windows-1251',
  'koi8-r',
  'windows-1250',
  'windows-1253',
  'windows-1255',
  'windows-1256',
  'windows-874',
] as const

export type SelectableEncoding = (typeof SELECTABLE_ENCODINGS)[number]

/** shell-wide AutoSave default; updatedAt is 0 until the user has ever set it */
export interface AutoSaveDefault {
  on: boolean
  updatedAt: number
}

// ---- chat attachments (local files fed to the agent via tools) ----

/** image attachments skip text extraction: read as base64 on send and passed to the model with the user message */
export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  path: string
  name: string
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  /** per-file reasons, already localized */
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  name?: string
  totalChars?: number
  offset?: number
  text?: string
  error?: string
}

export interface AttachmentImageResult {
  ok: boolean
  base64?: string
  mime?: string
  error?: string
}

export type SaveMode = 'save' | 'saveAs'

/**
 * Source-editor view preferences (UX-1704), persisted workspace-wide in the
 * shared userData/app-settings.json under one key. wordWrap keeps the
 * historical default (the editor always wrapped before the toggle);
 * scrollSync defaults off (an optional split-view affordance).
 */
export interface EditorPrefs {
  wordWrap: boolean
  scrollSync: boolean
}

/** a partial update; only the boolean-named fields are applied */
export type EditorPrefsPatch = Partial<EditorPrefs>

export const DEFAULT_EDITOR_PREFS: EditorPrefs = { wordWrap: true, scrollSync: false }

export interface SaveHtmlRequest {
  /** full document text (frontmatter included) */
  text: string
  /** Authored image paths in document order; the main process validates every path. */
  imageSources: string[]
  mode: SaveMode
  /**
   * Silent first save for an untitled document (AI auto-naming): saves to a
   * unique path under Documents derived from this name, without a dialog.
   * Ignored when the document already has a path.
   */
  suggestedName?: string
  /** Default file name offered by the save dialog of an untitled document */
  defaultName?: string
  /**
   * Marks an automatic save (the autosave tick / window blur). When the
   * staleness fence refuses the write, an auto save is declined without the
   * external-change dialog instead of popping a modal every 30 seconds.
   */
  auto?: boolean
}

export type SaveHtmlResult =
  | {
      ok: true
      path: string
      /** Save As may relocate local images into the new document's assets directory. */
      imageRewrites?: Array<{ from: string; to: string }>
    }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  webSearch: 'ai:web-search',
  imageSearch: 'ai:image-search',
} as const

export interface WebSearchResult {
  answer?: string
  results: Array<{ title: string; url: string; snippet: string }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export interface ImageSearchResult {
  images: Array<{ title?: string; imageUrl: string; width?: number; height?: number }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

export type ExportFormat = 'pdf' | 'docx'

/** Word export: html2docx renders the document in a hidden window and writes native OOXML; the result opens in Docs */
export interface ExportDocxRequest {
  /** the document text */
  html: string
  /** file name (no extension) suggested in the dialog */
  suggestedName: string
}

export interface ExportPdfRequest {
  /** self-contained print HTML */
  html: string
  suggestedName: string
}

export type ExportResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

export interface ImageData {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

/** API exposed by preload to the renderer (window.htmlApi) */
export interface HtmlApi {
  /** Take the md path pending for this view (queued at tab creation); null = new untitled document */
  consumePending(): Promise<string | null>
  /** Read the file as UTF-8 text. Only paths granted to this view are allowed */
  readFile(path: string): Promise<{ text: string; recovered: boolean }>
  /**
   * Remember (encoding) or forget (null, back to auto-detection) the charset
   * this file decodes with; the caller re-reads via readFile afterwards
   * (UX-1696 "Reopen with encoding"). Only paths granted to this view are
   * allowed; the pick survives tab close and relaunch (UX-1653).
   */
  setEncoding(path: string, encoding: string | null): Promise<boolean>
  /** Editor view preferences (word wrap, split scroll-sync), persisted workspace-wide (UX-1704) */
  getEditorPrefs(): Promise<EditorPrefs>
  /**
   * Merge a partial editor-prefs update through the app-settings single-writer
   * queue (PR #108 discipline: merge-write that keeps untouched keys);
   * resolves with the stored value.
   */
  setEditorPrefs(patch: EditorPrefsPatch): Promise<EditorPrefs>
  /** crash-recovery copy push (dirty renderers, every ~30s and on blur) */
  writeRecovery(path: string, text: string): Promise<void>
  /** Push the current buffer so html-preview:// serves it to the preview iframe */
  updatePreview(text: string): void
  /** The html-preview:// URL bound to this view (a present tab gets its owner's URL) */
  getPreviewInfo(): Promise<{ url: string }>
  /** Present → Fullscreen: cover the screen in one main-side call (tab-strip bleed, macOS simpleFullScreen) */
  setPresentFullScreen(on: boolean): Promise<void>
  /** Present → New tab: a chrome-free tab (shell) or window (standalone) showing this view's preview */
  presentInNewTab(title: string): Promise<boolean>
  /**
   * Write the document text. With a granted file path the write is atomic
   * (tmp + rename); untitled documents and mode 'saveAs' go through a main-process
   * save dialog first. The resolved path is granted to the view and returned.
   */
  save(request: SaveHtmlRequest): Promise<SaveHtmlResult>
  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */
  setDirty(dirty: boolean): void
  /** Shell menu Save / Save As → renderer serializes and calls save() with the given mode */
  onSaveRequest(handler: (mode: SaveMode) => void): () => void
  /** Resolves a menu-save waiter when doSave exits without ever invoking save() (busy/loading) */
  sendSaveRequestAck(ok: boolean): void
  /** Main process picked "Save" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  /** The file was renamed on disk (Home list rename) — renderer syncs its display path */
  onFileRenamed(handler: (newPath: string) => void): () => void
  /** Name an untitled document's tab before its first save (derived from the user's first AI request) */
  setProvisionalTitle(title: string): void
  /**
   * Pick an image file and copy it into `assets/` next to the open document;
   * returns the relative path to author into the html, or null when the
   * document is untitled or the picker was canceled.
   */
  pickImage(): Promise<string | null>
  /**
   * Persist pasted/dropped image bytes into `assets/` next to the open
   * document; returns the relative path to author, or null when untitled.
   */
  saveImage(data: { base64: string; ext: string }): Promise<string | null>
  /**
   * Read an image referenced by the document for DOCX embedding. Only paths
   * inside the document's directory are allowed; anything else returns null.
   */
  readImage(src: string): Promise<ImageData | null>
  /** file picker for chat attachments (multi-select) */
  pickAttachments(): Promise<AttachmentAddResult | null>
  /** validate dropped paths and return attachment metadata */
  addAttachmentPaths(paths: string[]): Promise<AttachmentAddResult>
  /** persist a pasted clipboard image (no local path) to a temp file and add it as an attachment */
  addPastedImage(data: ArrayBuffer, ext: string): Promise<AttachmentAddResult>
  /** read a slice of the extracted text of an attachment */
  readAttachment(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  /** read an image attachment as base64 for multimodal input */
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  /** local path of a dropped / pasted File, '' when it has none */
  getPathForFile(file: File): string
  /** Shell menu export → renderer serializes and calls exportDocx/exportPdf */
  onExportRequest(handler: (format: ExportFormat) => void): () => void
  /** Shell menu Print → renderer builds the print HTML and opens the system print dialog */
  onPrintRequest(handler: () => void): () => void
  exportDocx(request: ExportDocxRequest): Promise<ExportResult>
  exportPdf(request: ExportPdfRequest): Promise<ExportResult>
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  getAutoSaveDefault(): Promise<AutoSaveDefault>
  onAutoSaveDefaultChanged(handler: (value: AutoSaveDefault) => void): () => void
  /** AI panel text size + chat-input spellcheck (Settings → General in the shell) */
  getAiPanelPrefs(): Promise<AiPanelPrefs>
  onAiPanelPrefsChanged(handler: (prefs: AiPanelPrefs) => void): () => void
  /** press on the shell chrome (tab strip is a sibling WebContentsView whose
   *  clicks produce no DOM event here) — dismiss open popovers */
  onChromePressed(handler: () => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /** Main-process web search (Serper/DuckDuckGo via the shared ai:web-search handler) */
  webSearch(query: string, maxResults?: number): Promise<WebSearchResult>
  /** Main-process image search (shared ai:image-search handler) */
  imageSearch(query: string, maxResults?: number): Promise<ImageSearchResult>
  /** Download an image URL in the main process (CORS-free, scheme/target validated) */
  fetchImage(url: string): Promise<ImageData | null>
  /** AI image generation via the configured media provider (html-owned channel) */
  aiGenerateImage(op: { prompt: string; aspectRatio?: string }): Promise<{
    url?: string
    error?: string
  }>
}
