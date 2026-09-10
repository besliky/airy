// Read-only text session for legacy formats the engines cannot edit without
// an external converter (Phase 3b): today the .doc path without LibreOffice,
// backed by word-extractor through @genoffice/file-parse. The agent gets the
// full text plus an explicit editable:false limitation in the meta; closing
// is a no-op (no temp files, no sidecar session).
import { randomUUID } from 'node:crypto'
import { readFile, stat } from 'node:fs/promises'

import { countWords } from '../docx/session.js'
import { resolveConfined } from '../docx/paths.js'

export interface TextSessionMeta {
  readonly handle: string
  readonly kind: 'text'
  readonly path: string
  readonly fileName: string
  readonly format: string
  readonly converted: boolean
  readonly editable: boolean
  readonly warnings: readonly string[]
  readonly wordCount: number
  readonly charCount: number
}

export class TextSession {
  readonly handle: string
  readonly path: string
  readonly format: string
  private readonly text: string
  private readonly warning: string

  private constructor(handle: string, path: string, format: string, text: string, warning: string) {
    this.handle = handle
    this.path = path
    this.format = format
    this.text = text
    this.warning = warning
  }

  /**
   * Open a legacy document inside the workspace root and extract its text.
   * `extract` decodes the bytes (e.g. word-extractor for .doc).
   */
  static async open(
    rawPath: string,
    root: string | undefined,
    options: {
      format: string
      warning: string
      extract: (bytes: Uint8Array) => Promise<string>
    },
  ): Promise<TextSession> {
    const path = resolveConfined(rawPath, root)
    let bytes: Uint8Array
    try {
      bytes = new Uint8Array(await readFile(path))
      await stat(path)
    } catch (e) {
      throw new Error(`Cannot read "${path}": ${e instanceof Error ? e.message : String(e)}`, {
        cause: e,
      })
    }
    let text: string
    try {
      text = await options.extract(bytes)
    } catch (e) {
      throw new Error(
        `Cannot extract text from "${path}": ${e instanceof Error ? e.message : String(e)}`,
        { cause: e },
      )
    }
    return new TextSession(randomUUID(), path, options.format, text, options.warning)
  }

  meta(): TextSessionMeta {
    return {
      handle: this.handle,
      kind: 'text',
      path: this.path,
      fileName: this.path.split('/').pop() ?? this.path,
      format: this.format,
      converted: false,
      editable: false,
      warnings: [this.warning],
      wordCount: countWords(this.text),
      charCount: this.text.length,
    }
  }

  /** The extracted text (blocks/range selections do not apply to plain text). */
  readDocument(): string {
    const maxChars = 30_000
    const body =
      this.text.length > maxChars
        ? `${this.text.slice(0, maxChars)}\n…(output truncated at ${String(maxChars)} characters)`
        : this.text
    return [
      `Read-only text extracted from "${this.path.split('/').pop() ?? this.path}" (format .${this.format}, editable: false).`,
      this.warning,
      '',
      body,
    ].join('\n')
  }

  close(): Promise<string[]> {
    return Promise.resolve([])
  }
}
