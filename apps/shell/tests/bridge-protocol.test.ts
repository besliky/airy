import { describe, expect, it } from 'vitest'

import { createBridgeDispatcher } from '../src/main/bridge/dispatcher'
import {
  BRIDGE_PROTOCOL_VERSION,
  BridgeMethodError,
  createHandshakeGate,
  encodeResponse,
  NdjsonFramer,
  parseRequestLine,
} from '../src/main/bridge/protocol'

/**
 * Bridge protocol v1 (src/main/bridge/protocol.ts + dispatcher.ts): NDJSON
 * codec rules, handshake gating, and dispatcher error mapping. Pure Node —
 * no Electron anywhere near this file.
 */

describe('parseRequestLine', () => {
  it('parses a well-formed request', () => {
    const parsed = parseRequestLine('{"protocol_version":1,"method":"ping","params":{"a":1}}')
    expect(parsed).toEqual({
      ok: true,
      value: { protocol_version: 1, method: 'ping', params: { a: 1 } },
    })
  })

  it('accepts a request without params', () => {
    const parsed = parseRequestLine('{"protocol_version":1,"method":"ping"}')
    expect(parsed.ok).toBe(true)
    if (parsed.ok) expect(parsed.value.params).toBeUndefined()
  })

  it('tolerates surrounding whitespace and CRLF remnants', () => {
    const parsed = parseRequestLine('  {"protocol_version":1,"method":"ping"}  ')
    expect(parsed.ok).toBe(true)
  })

  it('rejects empty lines as invalid_request', () => {
    const parsed = parseRequestLine('   ')
    expect(parsed).toEqual({
      ok: false,
      error: { code: 'invalid_request', message: 'empty line' },
    })
  })

  it('rejects malformed JSON as invalid_request', () => {
    const parsed = parseRequestLine('{"protocol_version":1,"method":')
    expect(parsed.ok).toBe(false)
    if (!parsed.ok) expect(parsed.error.code).toBe('invalid_request')
  })

  it('rejects non-object JSON as invalid_request', () => {
    for (const line of ['[1,2]', '"ping"', '42']) {
      const parsed = parseRequestLine(line)
      expect(parsed.ok).toBe(false)
      if (!parsed.ok) expect(parsed.error.code).toBe('invalid_request')
    }
  })

  it('rejects a missing/odd protocol_version as invalid_request', () => {
    for (const line of [
      '{"method":"ping"}',
      '{"protocol_version":"1","method":"ping"}',
      '{"protocol_version":1.5,"method":"ping"}',
    ]) {
      const parsed = parseRequestLine(line)
      expect(parsed.ok).toBe(false)
      if (!parsed.ok) expect(parsed.error.code).toBe('invalid_request')
    }
  })

  it('reports version skew as unsupported_version', () => {
    const parsed = parseRequestLine('{"protocol_version":2,"method":"ping"}')
    expect(parsed).toEqual({
      ok: false,
      error: {
        code: 'unsupported_version',
        message: 'protocol_version 2 is not supported (server speaks 1)',
      },
    })
  })

  it('rejects a missing method as invalid_request', () => {
    const parsed = parseRequestLine('{"protocol_version":1}')
    expect(parsed.ok).toBe(false)
    if (!parsed.ok) expect(parsed.error.code).toBe('invalid_request')
  })

  it('rejects non-object params as invalid_request', () => {
    const parsed = parseRequestLine('{"protocol_version":1,"method":"ping","params":[1]}')
    expect(parsed.ok).toBe(false)
    if (!parsed.ok) expect(parsed.error.code).toBe('invalid_request')
  })
})

describe('encodeResponse', () => {
  it('round-trips results and errors', () => {
    expect(JSON.parse(encodeResponse({ ok: true, result: { pong: true } }))).toEqual({
      ok: true,
      result: { pong: true },
    })
    expect(
      JSON.parse(encodeResponse({ ok: false, error: { code: 'timeout', message: 'x' } })),
    ).toEqual({ ok: false, error: { code: 'timeout', message: 'x' } })
  })
})

describe('NdjsonFramer', () => {
  it('accumulates chunks split mid-line', () => {
    const framer = new NdjsonFramer()
    expect(framer.push('{"meth')).toEqual({ lines: [], overflow: false })
    expect(framer.push('od":"pi')).toEqual({ lines: [], overflow: false })
    const result = framer.push('ng"}\n')
    expect(result.lines).toEqual(['{"method":"ping"}'])
    expect(result.overflow).toBe(false)
  })

  it('returns several completed lines per chunk and keeps the tail buffered', () => {
    const framer = new NdjsonFramer()
    const result = framer.push('a\nb\r\nc\nd')
    expect(result.lines).toEqual(['a', 'b', 'c'])
    expect(framer.push('\n').lines).toEqual(['d'])
  })

  it('accepts buffers', () => {
    const framer = new NdjsonFramer()
    expect(framer.push(Buffer.from('x\n')).lines).toEqual(['x'])
  })

  it('flags a line over the cap as overflow', () => {
    const framer = new NdjsonFramer(8)
    const result = framer.push('0123456789abcdef')
    expect(result.overflow).toBe(true)
    expect(result.lines).toEqual([])
  })
})

describe('createHandshakeGate', () => {
  const TOKEN = 'secret-token'

  it('admits a correct hello and opens the session', () => {
    const gate = createHandshakeGate(TOKEN)
    expect(
      gate.admit({ protocol_version: 1, method: 'hello', params: { token: TOKEN } }),
    ).toBeNull()
    expect(gate.authorized).toBe(true)
  })

  it('rejects a first message that is not hello', () => {
    const gate = createHandshakeGate(TOKEN)
    expect(gate.admit({ protocol_version: 1, method: 'ping', params: {} })).toEqual({
      code: 'unauthorized',
      message: 'the first request must be hello with the bridge token',
    })
    expect(gate.authorized).toBe(false)
  })

  it('rejects a wrong token', () => {
    const gate = createHandshakeGate(TOKEN)
    expect(gate.admit({ protocol_version: 1, method: 'hello', params: { token: 'nope' } })).toEqual(
      { code: 'unauthorized', message: 'invalid bridge token' },
    )
  })

  it('rejects a missing token', () => {
    const gate = createHandshakeGate(TOKEN)
    expect(gate.admit({ protocol_version: 1, method: 'hello', params: {} })).toEqual({
      code: 'unauthorized',
      message: 'invalid bridge token',
    })
  })

  it('lets any other method through once authorized', () => {
    const gate = createHandshakeGate(TOKEN)
    gate.admit({ protocol_version: 1, method: 'hello', params: { token: TOKEN } })
    expect(gate.admit({ protocol_version: 1, method: 'ping', params: {} })).toBeNull()
  })

  it('rejects a second hello', () => {
    const gate = createHandshakeGate(TOKEN)
    gate.admit({ protocol_version: 1, method: 'hello', params: { token: TOKEN } })
    expect(gate.admit({ protocol_version: 1, method: 'hello', params: { token: TOKEN } })).toEqual({
      code: 'invalid_request',
      message: 'hello is only valid as the first message',
    })
  })
})

describe('createBridgeDispatcher', () => {
  it('returns handler results', async () => {
    const dispatcher = createBridgeDispatcher({
      ping: () => ({ pong: true }),
    })
    const response = await dispatcher.call({ protocol_version: 1, method: 'ping', params: {} })
    expect(response).toEqual({ ok: true, result: { pong: true } })
  })

  it('lists the supported methods on unknown_method', async () => {
    const dispatcher = createBridgeDispatcher({ ping: () => 1, list: () => 2 })
    const response = await dispatcher.call({ protocol_version: 1, method: 'frobnicate' })
    expect(response).toEqual({
      ok: false,
      error: {
        code: 'unknown_method',
        message: 'unknown method "frobnicate" (supported: list, ping)',
      },
    })
  })

  it('re-checks the protocol version', async () => {
    const dispatcher = createBridgeDispatcher({ ping: () => 1 })
    const response = await dispatcher.call({ protocol_version: 99, method: 'ping' })
    expect(response).toEqual({
      ok: false,
      error: {
        code: 'unsupported_version',
        message: `protocol_version 99 is not supported (server speaks ${BRIDGE_PROTOCOL_VERSION})`,
      },
    })
  })

  it('aborts the handler signal and answers timeout when the handler stalls', async () => {
    let aborted = false
    const dispatcher = createBridgeDispatcher(
      {
        stall: (_params, signal) =>
          new Promise((_resolve, reject) => {
            signal.addEventListener('abort', () => {
              aborted = true
              reject(new Error('aborted'))
            })
          }),
      },
      { timeoutMs: 30 },
    )
    const response = await dispatcher.call({ protocol_version: 1, method: 'stall' })
    expect(response).toEqual({
      ok: false,
      error: { code: 'timeout', message: 'timed out after 30ms' },
    })
    expect(aborted).toBe(true)
  })

  it('maps thrown BridgeMethodError to its code, other throws to internal', async () => {
    const dispatcher = createBridgeDispatcher({
      coded: () => {
        throw new BridgeMethodError('not_docs_tab', 'the active tab is not a docs document')
      },
      boom: () => {
        throw new Error('kaboom')
      },
    })
    expect(await dispatcher.call({ protocol_version: 1, method: 'coded' })).toEqual({
      ok: false,
      error: { code: 'not_docs_tab', message: 'the active tab is not a docs document' },
    })
    expect(await dispatcher.call({ protocol_version: 1, method: 'boom' })).toEqual({
      ok: false,
      error: { code: 'internal', message: 'kaboom' },
    })
  })
})
