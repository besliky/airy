// Airy live bridge wire protocol (client side): FIFO NDJSON over a Unix domain
// socket (linux/mac) or a Windows named pipe. Lightweight twin of
// apps/shell/src/main/bridge/protocol.ts — duplicated on purpose so the MCP
// package never imports app (Electron-adjacent) sources; keep the two modules
// in sync when the protocol changes. Pure Node, no Electron imports.
export const BRIDGE_PROTOCOL_VERSION = 1

// Mixed-version residual (tolerated while protocol_version stays 1): the
// undo params carry `ownTurnsOnly` so a rollback cannot revert another
// copilot client's turn, but an app from before that field existed accepts
// the call and ignores the flag — under version skew a live_apply_ops
// rollback may therefore undo the LAST turn regardless of owner. The server
// rejects unknown protocol_versions outright, so this is bounded to
// same-version (1) peers that merely predate the field; revisit by bumping
// the version once the flag is old enough to require.

export type BridgeErrorCode =
  | 'unauthorized'
  | 'unsupported_version'
  | 'unknown_method'
  | 'invalid_request'
  | 'not_docs_tab'
  | 'no_active_document'
  | 'tab_closed'
  | 'stale_document'
  | 'turn_owned_by_other'
  | 'invalid_params'
  | 'nothing_to_undo'
  | 'timeout'
  | 'internal'

const BRIDGE_ERROR_CODES: readonly BridgeErrorCode[] = [
  'unauthorized',
  'unsupported_version',
  'unknown_method',
  'invalid_request',
  'not_docs_tab',
  'no_active_document',
  'tab_closed',
  'stale_document',
  'turn_owned_by_other',
  'invalid_params',
  'nothing_to_undo',
  'timeout',
  'internal',
]

/** type guard for codes arriving on the wire (stringly-typed in the envelope) */
export function isBridgeErrorCode(value: string): value is BridgeErrorCode {
  return (BRIDGE_ERROR_CODES as readonly string[]).includes(value)
}

export interface BridgeRequest {
  protocol_version: number
  method: string
  params?: Record<string, unknown>
}

export interface BridgeError {
  code: BridgeErrorCode
  message: string
}

export type BridgeResponse = { ok: true; result: unknown } | { ok: false; error: BridgeError }

export interface FramerResult {
  /** completed lines, without their newline terminators */
  lines: string[]
  /** a single line exceeded the cap — the connection must be dropped */
  overflow: boolean
}

/**
 * Accumulates socket chunks into complete NDJSON lines. Carriage returns are
 * tolerated so a CRLF-flavored peer still frames cleanly (same contract as the
 * bridge server's framer).
 */
export class NdjsonFramer {
  private buffer = ''
  constructor(
    /** caps a single message so a rogue peer cannot grow memory without bound */
    readonly maxLineBytes = 8 * 1024 * 1024,
  ) {}

  push(chunk: string | Buffer): FramerResult {
    this.buffer += typeof chunk === 'string' ? chunk : chunk.toString('utf8')
    const lines: string[] = []
    let newline = this.buffer.indexOf('\n')
    while (newline !== -1) {
      const line = this.buffer.slice(0, newline).replace(/\r$/, '')
      this.buffer = this.buffer.slice(newline + 1)
      lines.push(line)
      newline = this.buffer.indexOf('\n')
    }
    const overflow = this.buffer.length > this.maxLineBytes
    return { lines, overflow }
  }
}

/**
 * Parse one NDJSON request line (the format the client sends and the test
 * mock server receives). Structural problems are `invalid_request`; a present
 * but unsupported protocol_version is `unsupported_version` so version-skew is
 * reported per message.
 */
export function parseRequestLine(
  line: string,
): { ok: true; value: BridgeRequest } | { ok: false; error: BridgeError } {
  const trimmed = line.trim()
  if (!trimmed) return { ok: false, error: { code: 'invalid_request', message: 'empty line' } }
  let parsed: unknown
  try {
    parsed = JSON.parse(trimmed)
  } catch {
    return { ok: false, error: { code: 'invalid_request', message: 'line is not valid JSON' } }
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    return {
      ok: false,
      error: { code: 'invalid_request', message: 'request must be a JSON object' },
    }
  }
  const obj = parsed as Record<string, unknown>
  const version = obj.protocol_version
  if (typeof version !== 'number' || !Number.isInteger(version)) {
    return {
      ok: false,
      error: { code: 'invalid_request', message: 'protocol_version must be an integer' },
    }
  }
  if (version !== BRIDGE_PROTOCOL_VERSION) {
    return {
      ok: false,
      error: {
        code: 'unsupported_version',
        message: `protocol_version ${version} is not supported (server speaks ${BRIDGE_PROTOCOL_VERSION})`,
      },
    }
  }
  if (typeof obj.method !== 'string' || obj.method.length === 0) {
    return {
      ok: false,
      error: { code: 'invalid_request', message: 'method must be a non-empty string' },
    }
  }
  if (
    obj.params !== undefined &&
    (typeof obj.params !== 'object' || obj.params === null || Array.isArray(obj.params))
  ) {
    return { ok: false, error: { code: 'invalid_request', message: 'params must be an object' } }
  }
  return {
    ok: true,
    value: {
      protocol_version: version,
      method: obj.method,
      ...(obj.params !== undefined ? { params: obj.params as Record<string, unknown> } : {}),
    },
  }
}

/**
 * Parse one NDJSON response line (what the bridge server replies). Structural
 * problems — bad JSON, a missing envelope — are reported as a message instead
 * of a typed error: the only sane client reaction is to drop the connection.
 */
export function parseResponseLine(
  line: string,
): { ok: true; value: BridgeResponse } | { ok: false; message: string } {
  const trimmed = line.trim()
  if (!trimmed) return { ok: false, message: 'empty response line' }
  let parsed: unknown
  try {
    parsed = JSON.parse(trimmed)
  } catch {
    return { ok: false, message: 'response line is not valid JSON' }
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    return { ok: false, message: 'response must be a JSON object' }
  }
  const obj = parsed as Record<string, unknown>
  if (obj.ok === true) {
    return { ok: true, value: { ok: true, result: obj.result } }
  }
  if (obj.ok === false) {
    const error = obj.error
    if (
      typeof error === 'object' &&
      error !== null &&
      !Array.isArray(error) &&
      typeof (error as Record<string, unknown>).code === 'string' &&
      typeof (error as Record<string, unknown>).message === 'string'
    ) {
      const { code, message } = error as { code: string; message: string }
      return {
        ok: true,
        value: {
          ok: false,
          error: { code: isBridgeErrorCode(code) ? code : 'internal', message },
        },
      }
    }
  }
  return { ok: false, message: 'response is not a valid {ok,result}|{ok,error} envelope' }
}
