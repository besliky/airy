import { afterEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readAiSettingsFile } from '../src/media-tools'
import { setAiSecretDecrypter } from '@airy-office/ai-provider'

afterEach(() => {
  setAiSecretDecrypter(undefined)
})

describe('readAiSettingsFile secret decoding', () => {
  it('decrypts enc: keys through the registered decrypter', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'airy-ai-settings-'))
    try {
      const encrypted = Buffer.from('serper-decrypted-key', 'utf-8').toString('base64')
      await writeFile(
        join(dir, 'ai-settings.json'),
        JSON.stringify({
          v: 1,
          provider: 'none',
          providers: {},
          search: {
            provider: 'serper',
            providers: { serper: { apiKey: `enc:${encrypted}` }, tavily: { apiKey: '' } },
          },
        }),
      )
      setAiSecretDecrypter((blob) => Buffer.from(blob, 'base64').toString('utf-8'))
      const settings = readAiSettingsFile(join(dir, 'ai-settings.json'))
      expect(settings.search?.providers.serper.apiKey).toBe('serper-decrypted-key')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })

  it('reads legacy plaintext files unchanged and empty enc: without a decrypter', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'airy-ai-settings-'))
    try {
      await writeFile(
        join(dir, 'ai-settings.json'),
        JSON.stringify({
          v: 1,
          provider: 'none',
          providers: {},
          search: {
            provider: 'tavily',
            providers: { serper: { apiKey: 'enc:whatever' }, tavily: { apiKey: 'plain-tavily' } },
          },
        }),
      )
      // no decrypter registered: enc: reads as an unavailable key, plaintext survives
      const settings = readAiSettingsFile(join(dir, 'ai-settings.json'))
      expect(settings.search?.providers.serper.apiKey).toBe('')
      expect(settings.search?.providers.tavily.apiKey).toBe('plain-tavily')
      expect(settings.search?.provider).toBe('tavily')
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  })
})
