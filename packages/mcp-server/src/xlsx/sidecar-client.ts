// Headless client for the xlsx sidecar (apps/sheets/native/xlsx-engine): a
// Rust process speaking line-delimited JSON over stdio. This is the MCP
// server's port of the app client (apps/sheets/src/main/xlsx-sidecar-client.ts,
// same wire protocol) — kept protocol-compatible but tuned for the headless
// use case: 'en' open locale by default (the app defaults to 'zh'), 30 s
// request timeouts with 120 s for archive-class commands (convert/save).
//
// Protocol (SA4): every request is {version:1, requestId, command, ...fields}
// on one line; replies are {version, requestId, ok, result|error}. requestId
// correlates replies (recalc replies may arrive out of order). Non-JSON or
// wrong-version lines on stdout are treated as library noise (IronCalc prints
// diagnostics there) and skipped, exactly like the app client.
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { createInterface, type Interface } from 'node:readline'

const PROTOCOL_VERSION = 1
const REQUEST_TIMEOUT_MS = 30_000
// Whole-archive commands (convert, zip rebuild) stream entire workbooks.
const ARCHIVE_TIMEOUT_MS = 120_000
const MAX_STDERR_LENGTH = 8_192

export interface SidecarError {
  readonly code: string
  readonly message: string
}

export interface SidecarResponse {
  readonly version: number
  readonly requestId: string
  readonly ok: boolean
  readonly result?: unknown
  readonly error?: SidecarError
}

interface PendingRequest {
  readonly resolve: (value: unknown) => void
  readonly reject: (error: Error) => void
  readonly timeout: NodeJS.Timeout
  readonly command: string
  /** Session-scoped commands record it so close() can sweep their queue slots. */
  readonly sessionId?: string
}

/** Pure reads a departing consumer can abandon; skipping a queued close (or
 * save) instead would leave its side effects permanently undone. */
const CANCELLABLE_READS = new Set(['read_range', 'read_formula_cells', 'read_media'])

/**
 * Encode one NDJSON request line. Pure so the protocol codec can be unit
 * tested without a process.
 */
export function encodeRequest(
  requestId: string,
  command: Readonly<Record<string, unknown>>,
): string {
  return `${JSON.stringify({ version: PROTOCOL_VERSION, requestId, ...command })}\n`
}

/**
 * Decode one stdout line into a protocol reply, or null when the line is
 * noise (unparseable JSON, wrong version, malformed envelope). Pure.
 */
export function parseSidecarLine(line: string): SidecarResponse | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(line)
  } catch {
    return null
  }
  if (typeof parsed !== 'object' || parsed === null) return null
  const response = parsed as Partial<SidecarResponse>
  if (
    response.version !== PROTOCOL_VERSION ||
    typeof response.requestId !== 'string' ||
    typeof response.ok !== 'boolean'
  ) {
    return null
  }
  return response as SidecarResponse
}

/** The sidecar surface the xlsx session and the save gateway need. */
export interface XlsxIo {
  open(path: string, locale?: string, shortDateFormat?: string): Promise<unknown>
  readRange(input: {
    readonly sessionId: string
    readonly sheetId: string
    readonly range: {
      readonly startRow: number
      readonly endRow: number
      readonly startColumn: number
      readonly endColumn: number
    }
  }): Promise<unknown>
  readFormulaCells(input: {
    readonly sessionId: string
    readonly sheetId: string
  }): Promise<unknown>
  recalcCells(input: {
    readonly path: string
    readonly edits: readonly {
      readonly sheet: string
      readonly row: number
      readonly column: number
      readonly input: string
    }[]
    readonly reads: readonly {
      readonly sheet: string
      readonly range: {
        readonly startRow: number
        readonly endRow: number
        readonly startColumn: number
        readonly endColumn: number
      }
    }[]
  }): Promise<unknown>
  /**
   * PERF-1778: after an in-place save whose recalc overlay just succeeded,
   * refresh the sidecar's resident recalc model to the rewritten file so
   * the next save's overlay is a resident hit instead of a whole-book
   * re-import. Sidecars predating the command answer a request error —
   * callers treat this as best-effort and keep the rebuild behavior.
   */
  restampRecalc(path: string): Promise<unknown>
  close(sessionId: string): Promise<void>
  convertWorkbook(input: { readonly path: string; readonly targetPath: string }): Promise<unknown>
  archiveManifest(path: string): Promise<unknown>
  readEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly outputDir: string
  }): Promise<unknown>
  scanEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly needle: string
  }): Promise<unknown>
  saveArchive(input: {
    readonly sourcePath: string
    readonly targetPath: string
    readonly replacements: readonly { name: string; contentPath: string }[]
    readonly removals: readonly string[]
    readonly additions: readonly { name: string; contentPath: string }[]
  }): Promise<unknown>
}

export class XlsxSidecarClient implements XlsxIo {
  private process: ChildProcessWithoutNullStreams | null = null
  private lines: Interface | null = null
  private readonly pending = new Map<string, PendingRequest>()
  private stderr = ''

  constructor(private readonly binaryPath: string) {}

  async open(path: string, locale = 'en', shortDateFormat?: string): Promise<unknown> {
    return this.request({
      command: 'open',
      path,
      locale,
      ...(shortDateFormat === undefined ? {} : { shortDateFormat }),
    })
  }

  async readRange(input: {
    readonly sessionId: string
    readonly sheetId: string
    readonly range: {
      readonly startRow: number
      readonly endRow: number
      readonly startColumn: number
      readonly endColumn: number
    }
  }): Promise<unknown> {
    return this.request({ command: 'read_range', ...input })
  }

  async readFormulaCells(input: {
    readonly sessionId: string
    readonly sheetId: string
  }): Promise<unknown> {
    return this.request({ command: 'read_formula_cells', ...input })
  }

  async recalcCells(input: {
    readonly path: string
    readonly edits: readonly {
      readonly sheet: string
      readonly row: number
      readonly column: number
      readonly input: string
    }[]
    readonly reads: readonly {
      readonly sheet: string
      readonly range: {
        readonly startRow: number
        readonly endRow: number
        readonly startColumn: number
        readonly endColumn: number
      }
    }[]
  }): Promise<unknown> {
    // the recalc worker runs the IronCalc import + evaluation off the request
    // loop; a cold import alone can take seconds, so give it the archive budget
    return this.request({ command: 'recalc_cells', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async restampRecalc(path: string): Promise<unknown> {
    // a metadata read + stamp write: plain request budget is plenty
    return this.request({ command: 'restamp_recalc', path })
  }

  async close(sessionId: string): Promise<void> {
    // Skip the session's queued reads so close does not wait behind them;
    // each settles with the sidecar's authoritative "cancelled" failure.
    for (const [requestId, pending] of this.pending) {
      if (pending.sessionId === sessionId && CANCELLABLE_READS.has(pending.command)) {
        this.sendRaw(
          encodeRequest(crypto.randomUUID(), { command: 'cancel', targetRequestId: requestId }),
        )
      }
    }
    await this.request({ command: 'close', sessionId })
  }

  async convertWorkbook(input: { path: string; targetPath: string }): Promise<unknown> {
    return this.request({ command: 'convert_workbook', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async archiveManifest(path: string): Promise<unknown> {
    return this.request({ command: 'archive_manifest', path }, ARCHIVE_TIMEOUT_MS)
  }

  async readEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly outputDir: string
  }): Promise<unknown> {
    return this.request({ command: 'read_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async scanEntries(input: {
    readonly path: string
    readonly entries: readonly string[]
    readonly needle: string
  }): Promise<unknown> {
    return this.request({ command: 'scan_entries', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  async saveArchive(input: {
    readonly sourcePath: string
    readonly targetPath: string
    readonly replacements: readonly { name: string; contentPath: string }[]
    readonly removals: readonly string[]
    readonly additions: readonly { name: string; contentPath: string }[]
  }): Promise<unknown> {
    return this.request({ command: 'save_archive', ...input }, ARCHIVE_TIMEOUT_MS)
  }

  getProcessId(): number | null {
    return this.process?.pid ?? null
  }

  /** Kill the sidecar process and fail every pending request. */
  stop(): void {
    this.lines?.close()
    this.lines = null
    this.process?.kill()
    this.process = null
    this.rejectPending(new Error('XLSX sidecar stopped.'))
  }

  /**
   * Fire-and-forget cancel: handled out of band by the sidecar's reader
   * thread, so it takes effect while earlier requests still wait in the
   * queue. The reply matches no pending entry and is dropped.
   */
  private sendRaw(payload: string): void {
    const child = this.process
    if (!child || child.killed) return
    child.stdin.write(payload)
  }

  private request(
    command: Readonly<Record<string, unknown>>,
    timeoutMs: number = REQUEST_TIMEOUT_MS,
  ): Promise<unknown> {
    const child = this.ensureStarted()
    const requestId = crypto.randomUUID()
    const payload = encodeRequest(requestId, command)
    const commandName = command.command
    const sessionId = typeof command.sessionId === 'string' ? command.sessionId : undefined
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(requestId)
        // Still queued sidecar-side, the request would execute for a reply
        // nobody consumes (a 100k-cell range serializes ~5MB of JSON). A
        // close is the exception: its side effect IS the cleanup, and no
        // caller retries a timed-out close — skipping it would leak the
        // session's index threads and cache for good.
        if (commandName !== 'close') {
          this.sendRaw(
            encodeRequest(crypto.randomUUID(), { command: 'cancel', targetRequestId: requestId }),
          )
        }
        reject(new Error('XLSX sidecar request timed out.'))
      }, timeoutMs)
      this.pending.set(requestId, {
        resolve,
        reject,
        timeout,
        command: String(commandName),
        ...(sessionId === undefined ? {} : { sessionId }),
      })
      child.stdin.write(payload, (error) => {
        if (!error) return
        const pending = this.pending.get(requestId)
        if (!pending) return
        clearTimeout(pending.timeout)
        this.pending.delete(requestId)
        pending.reject(error)
      })
    })
  }

  private ensureStarted(): ChildProcessWithoutNullStreams {
    if (this.process && !this.process.killed) return this.process
    let child: ChildProcessWithoutNullStreams
    try {
      child = spawn(this.binaryPath, [], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
      })
    } catch (error) {
      // Synchronous spawn failures ("spawn UNKNOWN") carry no path in the
      // message — rethrow with enough context to diagnose from a log line.
      throw new Error(
        `XLSX sidecar failed to start (${this.binaryPath}): ${error instanceof Error ? error.message : String(error)}`,
        { cause: error },
      )
    }
    this.process = child
    this.stderr = ''
    this.lines = createInterface({ input: child.stdout })
    this.lines.on('line', (line) => this.handleLine(line))
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk: string) => {
      this.stderr = `${this.stderr}${chunk}`.slice(-MAX_STDERR_LENGTH)
    })
    child.once('error', (error) => {
      this.process = null
      this.rejectPending(error)
    })
    child.once('exit', (code, signal) => {
      this.process = null
      this.lines?.close()
      this.lines = null
      const detail = this.stderr.trim()
      const reason = detail
        ? `XLSX sidecar exited: ${detail}`
        : `XLSX sidecar exited with code ${String(code)} and signal ${String(signal)}.`
      this.rejectPending(new Error(reason))
    })
    return child
  }

  private handleLine(line: string): void {
    const response = parseSidecarLine(line)
    if (response === null) {
      // Library code inside the sidecar prints diagnostics to stdout
      // (IronCalc's importer does). A genuinely garbled response surfaces as
      // that one request's timeout instead.
      this.stderr = `${this.stderr}[stdout] ${line}\n`.slice(-MAX_STDERR_LENGTH)
      return
    }
    const pending = this.pending.get(response.requestId)
    if (!pending) return
    clearTimeout(pending.timeout)
    this.pending.delete(response.requestId)
    if (response.ok) {
      pending.resolve(response.result)
      return
    }
    pending.reject(new Error(response.error?.message ?? 'XLSX sidecar request failed.'))
  }

  private rejectPending(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timeout)
      pending.reject(error)
    }
    this.pending.clear()
  }
}
