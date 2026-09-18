// Airy live bridge protocol v1: FIFO NDJSON over a Unix domain socket
// (linux/mac) or a Windows named pipe. Pure Node module — no Electron imports
// — so the codec, framer and handshake rules are unit-testable in plain Node
// (tests/bridge-protocol.test.ts). The xlsx sidecar uses the same
// version-in-every-message discipline: protocol skew is caught per request,
// not just once at handshake time.
import { timingSafeEqual } from 'node:crypto'

export const BRIDGE_PROTOCOL_VERSION = 1

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

export function bridgeError(code: BridgeErrorCode, message: string): BridgeResponse {
  return { ok: false, error: { code, message } }
}

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

/** type guard for codes arriving from renderer-side handlers (stringly-typed on the wire) */
export function isBridgeErrorCode(value: string): value is BridgeErrorCode {
  return (BRIDGE_ERROR_CODES as readonly string[]).includes(value)
}

/**
 * A handler-side failure carrying a protocol error code. The dispatcher maps
 * these onto {ok:false,error} responses; anything else thrown becomes
 * `internal`, keeping client bugs out of the wire contract.
 */
export class BridgeMethodError extends Error {
  readonly code: BridgeErrorCode
  constructor(code: BridgeErrorCode, message: string) {
    super(message)
    this.name = 'BridgeMethodError'
    this.code = code
  }
}

// ---- NDJSON codec ----

/**
 * Parse one NDJSON request line. Structural problems (bad JSON, missing
 * method) are `invalid_request`; a present but unsupported protocol_version is
 * `unsupported_version` so version-skew is reported per message.
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

export function encodeResponse(response: BridgeResponse): string {
  return JSON.stringify(response)
}

// ---- NDJSON framing ----

export interface FramerResult {
  /** completed lines, without their newline terminators */
  lines: string[]
  /** a single line exceeded the cap — the connection must be dropped */
  overflow: boolean
}

/**
 * Arrivals below this size coalesce into the framer's tail scratch (see
 * NdjsonFramer.append): a dribbled line arrives as many tiny chunks, and each
 * Buffer object costs ~an order of magnitude more than the byte it carries.
 */
const COALESCE_LIMIT_BYTES = 64 * 1024
/** initial capacity of a fresh tail scratch — doubling keeps appends amortized */
const SCRATCH_MIN_BYTES = 4 * 1024

/**
 * Accumulates socket chunks into complete NDJSON lines. Carriage returns are
 * tolerated so a CRLF-flavored client (or terminal echo) still frames cleanly.
 *
 * Input is accumulated as raw bytes and only complete lines are decoded: a
 * socket chunk may split a multi-byte UTF-8 sequence mid-character, and
 * decoding each chunk separately would turn the halves into U+FFFD.
 *
 * The pending tail is kept as the list of chunks it arrived in: rebuilding a
 * single buffer per push made a dribbled unterminated line O(n²) in copied
 * bytes, and only the extracted lines are ever concatenated. Dribbled small
 * arrivals are additionally coalesced into that last chunk (a scratch buffer
 * with doubling headroom), so a slow-dripped line also costs one live Buffer
 * instead of one per socket chunk (BUG-903).
 */
export class NdjsonFramer {
  private chunks: Buffer[] = []
  private pending = 0
  /** resume point of the LF scan, right after the last extracted line */
  private scanChunk = 0
  private scanOffset = 0
  /** backing store of the coalescing tail; chunks[chunks.length-1] is its used-prefix view */
  private scratch: Buffer | null = null
  private scratchCap = 0
  constructor(
    /** caps a single message so a rogue client cannot grow memory without bound */
    readonly maxLineBytes = 8 * 1024 * 1024,
  ) {}

  push(chunk: string | Buffer): FramerResult {
    const incoming = typeof chunk === 'string' ? Buffer.from(chunk, 'utf8') : chunk
    if (incoming.length > 0) {
      this.append(incoming)
      this.pending += incoming.length
    }
    const lines: string[] = []
    // 0x0A never appears inside a multi-byte UTF-8 sequence (continuation
    // bytes are >= 0x80), so line boundaries are safe to find on raw bytes.
    // The scan starts after the last extracted LF: earlier bytes hold none.
    for (;;) {
      let foundChunk = -1
      let foundAt = -1
      for (let index = this.scanChunk; index < this.chunks.length; index += 1) {
        const from = index === this.scanChunk ? this.scanOffset : 0
        const at = this.chunks[index]!.indexOf(0x0a, from)
        if (at !== -1) {
          foundChunk = index
          foundAt = at
          break
        }
      }
      if (foundChunk === -1) break
      // a 0x0D directly before the LF is CRLF noise, never a sequence byte;
      // it can sit at the end of the previous chunk when the LF opens this one
      const found = this.chunks[foundChunk]!
      let crTrim = 0
      if (foundAt > 0) {
        if (found[foundAt - 1] === 0x0d) crTrim = 1
      } else if (foundChunk > 0) {
        const previous = this.chunks[foundChunk - 1]!
        if (previous[previous.length - 1] === 0x0d) crTrim = 1
      }
      // the line = full earlier chunks + this chunk's prefix, minus the CR;
      // when the CR closed the previous chunk, that chunk joins as a partial
      let headParts: Buffer[]
      if (foundChunk === 0) {
        headParts = []
      } else if (crTrim === 1 && foundAt === 0) {
        const previous = this.chunks[foundChunk - 1]!
        const kept = previous.subarray(0, previous.length - 1)
        headParts = [...this.chunks.slice(0, foundChunk - 1), ...(kept.length > 0 ? [kept] : [])]
      } else {
        headParts = this.chunks.slice(0, foundChunk)
      }
      const tailPart = found.subarray(0, foundAt - (crTrim === 1 && foundAt > 0 ? 1 : 0))
      let lineBytes = tailPart.length
      for (const part of headParts) lineBytes += part.length
      // the cap binds completed lines too, not just the pending tail: a
      // push() caller may hand over one arbitrarily large buffer
      if (lineBytes > this.maxLineBytes) return { lines, overflow: true }
      lines.push(
        (headParts.length === 0 ? tailPart : Buffer.concat([...headParts, tailPart])).toString(
          'utf8',
        ),
      )
      // consume everything through the LF; fully drained chunks drop off
      this.pending -= lineBytes + crTrim + 1
      const rest = found.subarray(foundAt + 1)
      if (foundChunk > 0) this.chunks.splice(0, foundChunk)
      if (rest.length === 0) this.chunks.shift()
      else this.chunks[0] = rest
      this.scanChunk = 0
      this.scanOffset = 0
      // extraction may have consumed into the tail scratch — its live prefix
      // moved, so freeze it; the next small arrival starts a fresh scratch
      this.scratch = null
    }
    if (this.chunks.length === 0) {
      this.scanChunk = 0
      this.scanOffset = 0
    } else {
      this.scanChunk = this.chunks.length - 1
      this.scanOffset = this.chunks[this.chunks.length - 1]!.length
    }
    // the tail without a newline stays as bytes until more chunks arrive (a
    // peer that never terminates it simply never gets the line); the cap
    // counts bytes of that pending tail
    const overflow = this.pending > this.maxLineBytes
    return { lines, overflow }
  }

  /**
   * Store one arrival. Dribbled small chunks (below COALESCE_LIMIT_BYTES, no
   * newline of their own) are memcpy'd into a tail scratch buffer with
   * doubling headroom — one live Buffer for the whole dribble, each byte
   * copied O(1) times amortized. Concat-per-push would reintroduce the O(n²)
   * tail, and keeping every tiny chunk as its own object amplifies memory
   * ~100x: an 8MB line dribbled one byte at a time is millions of Buffer
   * objects while its bytes stay within the cap (BUG-903). Chunks that carry
   * their own newline never accumulate for long, so they bypass the scratch
   * and stay stored by reference like everything larger.
   */
  private append(incoming: Buffer): void {
    const view = this.chunks[this.chunks.length - 1]
    const scratchable = incoming.length < COALESCE_LIMIT_BYTES && !incoming.includes(0x0a)
    if (this.scratch !== null && view !== undefined && scratchable) {
      const used = view.length
      if (used + incoming.length <= this.scratchCap) {
        incoming.copy(this.scratch, used)
        this.chunks[this.chunks.length - 1] = this.scratch.subarray(0, used + incoming.length)
        return
      }
      let cap = this.scratchCap * 2
      while (cap < used + incoming.length) cap *= 2
      const grown = Buffer.allocUnsafe(cap)
      this.scratch.copy(grown, 0, 0, used)
      incoming.copy(grown, used)
      this.scratch = grown
      this.scratchCap = cap
      this.chunks[this.chunks.length - 1] = grown.subarray(0, used + incoming.length)
      return
    }
    if (scratchable) {
      let cap = SCRATCH_MIN_BYTES
      while (cap < incoming.length) cap *= 2
      const started = Buffer.allocUnsafe(cap)
      incoming.copy(started, 0)
      this.scratch = started
      this.scratchCap = cap
      this.chunks.push(started.subarray(0, incoming.length))
      return
    }
    // a large or self-terminated arrival: store by reference, freeze any scratch
    this.scratch = null
    this.chunks.push(incoming)
  }
}

// ---- handshake ----

function tokensMatch(a: string, b: string): boolean {
  const ba = Buffer.from(a, 'utf8')
  const bb = Buffer.from(b, 'utf8')
  if (ba.length !== bb.length) return false
  return timingSafeEqual(ba, bb)
}

/**
 * Per-connection handshake state: the first request on a fresh connection must
 * be `hello` carrying the token from the 0600 info file. Anything else — or a
 * wrong token — rejects with `unauthorized` and the connection is closed.
 */
export function createHandshakeGate(token: string): {
  /** null = admitted (or already authorized); an error also closes the connection */
  admit(request: BridgeRequest): BridgeError | null
  readonly authorized: boolean
} {
  let authorized = false
  return {
    get authorized() {
      return authorized
    },
    admit(request) {
      if (authorized) {
        return request.method === 'hello'
          ? { code: 'invalid_request', message: 'hello is only valid as the first message' }
          : null
      }
      if (request.method !== 'hello') {
        return {
          code: 'unauthorized',
          message: 'the first request must be hello with the bridge token',
        }
      }
      const presented = (request.params as { token?: unknown } | undefined)?.token
      if (typeof presented !== 'string' || !tokensMatch(presented, token)) {
        return { code: 'unauthorized', message: 'invalid bridge token' }
      }
      authorized = true
      return null
    },
  }
}
