import { describe, expect, it, vi } from 'vitest'
import { crashErrorPageUrl, isRecoverableRendererCrash, voidLoad } from '../src/process-safety'

describe('isRecoverableRendererCrash', () => {
  it('triggers recovery only for oom and crashed', () => {
    expect(isRecoverableRendererCrash('oom')).toBe(true)
    expect(isRecoverableRendererCrash('crashed')).toBe(true)
    // intentional teardown or launch failures that already surface elsewhere
    expect(isRecoverableRendererCrash('clean-exit')).toBe(false)
    expect(isRecoverableRendererCrash('killed')).toBe(false)
    expect(isRecoverableRendererCrash('launch-failed')).toBe(false)
    expect(isRecoverableRendererCrash('integrity-failure')).toBe(false)
  })
})

describe('voidLoad', () => {
  it('logs a rejected load instead of leaving it unhandled', async () => {
    const error = new Error('file not found')
    const spy = vi.spyOn(console, 'error').mockImplementation(() => undefined)
    voidLoad(Promise.reject(error), 'test renderer')
    // let the microtask queue drain so the catch ran
    await Promise.resolve()
    await Promise.resolve()
    expect(spy).toHaveBeenCalledWith('[load] test renderer failed:', error)
    spy.mockRestore()
  })

  it('stays silent for a successful load', async () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => undefined)
    voidLoad(Promise.resolve(), 'test renderer')
    await Promise.resolve()
    expect(spy).not.toHaveBeenCalled()
    spy.mockRestore()
  })
})

describe('crashErrorPageUrl', () => {
  it('builds a data URL and escapes markup in the message', () => {
    const url = crashErrorPageUrl('Tab <stopped> & needs attention')
    expect(url.startsWith('data:text/html;charset=utf-8,')).toBe(true)
    const decoded = decodeURIComponent(url.slice('data:text/html;charset=utf-8,'.length))
    expect(decoded).toContain('Tab &lt;stopped&gt; &amp; needs attention')
    expect(decoded).not.toContain('<stopped>')
  })
})
