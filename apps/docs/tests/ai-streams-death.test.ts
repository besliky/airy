import { EventEmitter } from 'node:events'
import { describe, expect, it } from 'vitest'

import { RendererAiStreams } from '../src/main/ai-streams'

/**
 * BUG-1698: the app-wide ai:stream registry must abort a renderer's in-flight
 * streams when that renderer dies mid-turn (render-process-gone: crash, oom,
 * kill -9 — the webContents object itself survives a hard kill for a while,
 * so 'destroyed' alone would fire too late) or is torn down gracefully
 * ('destroyed', the sheets BUG-407 semantics). A provider turn must observe
 * the abort signal, streams of other senders must be untouched, per-stream
 * cancel stays targeted, and the death hooks detach once a sender settles.
 */

class FakeSender extends EventEmitter {
  constructor(readonly id: number) {
    super()
  }
}

/** Provider-mock turn: rejects exactly when the stream's signal aborts. */
function providerTurn(signal: AbortSignal): Promise<never> {
  return new Promise((_resolve, reject) => {
    const onAbort = (): void => reject(new Error('provider request aborted'))
    if (signal.aborted) onAbort()
    else signal.addEventListener('abort', onAbort, { once: true })
  })
}

describe('RendererAiStreams', () => {
  it('render-process-gone aborts every in-flight stream of the dead renderer only', () => {
    const streams = new RendererAiStreams()
    const dead = new FakeSender(7)
    const first = streams.start(dead, 'req-1')
    const second = streams.start(dead, 'req-2')
    const other = streams.start(new FakeSender(8), 'req-3')

    dead.emit('render-process-gone', { reason: 'killed' })

    expect(first.signal.aborted).toBe(true)
    expect(second.signal.aborted).toBe(true)
    expect(other.signal.aborted).toBe(false)
  })

  it('a provider turn observes the abort raised by the renderer death', async () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(1)
    const controller = streams.start(sender, 'req-1')
    const turn = providerTurn(controller.signal)

    sender.emit('render-process-gone', { reason: 'killed' })

    await expect(turn).rejects.toThrow('provider request aborted')
  })

  it('destroyed webContents (graceful close) aborts its in-flight streams too', () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(2)
    const controller = streams.start(sender, 'req-1')

    sender.emit('destroyed')

    expect(controller.signal.aborted).toBe(true)
  })

  it('settled streams are untouched: a death after finish is a no-op', () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(3)
    const settled = streams.start(sender, 'req-1')
    streams.finish(sender, 'req-1')

    sender.emit('render-process-gone', { reason: 'killed' })
    sender.emit('destroyed')

    expect(settled.signal.aborted).toBe(false)
  })

  it('cancel aborts only the targeted stream and the sender stays wired', () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(4)
    const cancelled = streams.start(sender, 'req-1')
    const alive = streams.start(sender, 'req-2')

    streams.cancel('req-1')
    expect(cancelled.signal.aborted).toBe(true)
    expect(alive.signal.aborted).toBe(false)

    streams.finish(sender, 'req-1')
    sender.emit('render-process-gone', { reason: 'killed' })
    expect(alive.signal.aborted).toBe(true)
  })

  it('cancel of an unknown id is a no-op', () => {
    const streams = new RendererAiStreams()
    expect(() => streams.cancel('missing')).not.toThrow()
  })

  it('death hooks detach once the sender has no in-flight streams left', () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(5)
    const controller = streams.start(sender, 'req-1')

    expect(sender.listenerCount('destroyed')).toBe(1)
    expect(sender.listenerCount('render-process-gone')).toBe(1)

    streams.finish(sender, 'req-1')

    expect(sender.listenerCount('destroyed')).toBe(0)
    expect(sender.listenerCount('render-process-gone')).toBe(0)
    expect(controller.signal.aborted).toBe(false)
  })

  it('sequential streams of one sender keep a single pair of death hooks', () => {
    const streams = new RendererAiStreams()
    const sender = new FakeSender(6)
    const first = streams.start(sender, 'req-1')
    streams.finish(sender, 'req-1')
    const second = streams.start(sender, 'req-2')

    expect(sender.listenerCount('destroyed')).toBe(1)
    expect(sender.listenerCount('render-process-gone')).toBe(1)
    expect(first.signal.aborted).toBe(false)
    expect(second.signal.aborted).toBe(false)
  })
})
