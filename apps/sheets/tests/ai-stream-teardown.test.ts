/**
 * BUG-407: the AI stream registry must tie each tab's in-flight streams to
 * the tab lifecycle — closing the tab aborts every running provider request
 * so a dead renderer stops accumulating stream chunks. Settled streams must
 * be untouched by teardown, and a per-stream cancel must not disturb others.
 */
import { describe, expect, it } from 'vitest'
import { AiStreamRegistry } from '../src/main/ai-streams'

describe('AiStreamRegistry', () => {
  it('abortAll aborts every in-flight stream (tab teardown)', () => {
    const registry = new AiStreamRegistry()
    const first = registry.start('req-1')
    const second = registry.start('req-2')

    expect(first.signal.aborted).toBe(false)
    registry.abortAll()

    expect(first.signal.aborted).toBe(true)
    expect(second.signal.aborted).toBe(true)
  })

  it('cancel aborts only the targeted stream', () => {
    const registry = new AiStreamRegistry()
    const first = registry.start('req-1')
    const second = registry.start('req-2')

    registry.cancel('req-1')

    expect(first.signal.aborted).toBe(true)
    expect(second.signal.aborted).toBe(false)
  })

  it('finish drops the entry so teardown no longer touches the stream', () => {
    const registry = new AiStreamRegistry()
    const settled = registry.start('req-1')
    registry.finish('req-1')

    registry.abortAll()

    expect(settled.signal.aborted).toBe(false)
  })

  it('cancel of an unknown or finished id is a no-op', () => {
    const registry = new AiStreamRegistry()
    expect(() => registry.cancel('missing')).not.toThrow()

    const stream = registry.start('req-1')
    registry.finish('req-1')
    registry.cancel('req-1')
    expect(stream.signal.aborted).toBe(false)
  })

  it('abortAll clears the registry so later finish calls stay no-ops', () => {
    const registry = new AiStreamRegistry()
    const stream = registry.start('req-1')
    registry.abortAll()

    registry.finish('req-1')
    registry.cancel('req-1')
    expect(stream.signal.aborted).toBe(true)
  })
})
