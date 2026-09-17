import { afterEach, describe, expect, it } from 'vitest'
import {
  ENCRYPTED_SECRET_PREFIX,
  MASKED_SECRET_PLACEHOLDER,
  decryptStoredAiSettings,
  encryptStoredAiSettings,
  isEncryptedSecret,
  isMaskedSecret,
  maskAiSettingsSecrets,
  maskSecret,
  overlayAiSettingsSecrets,
  overlaySecret,
  setAiSecretDecrypter,
} from '../src/settings-secrets'
import { defaultAiSettings } from '../src/providers'
import type { AiSettings } from '../src/types'

afterEach(() => {
  setAiSecretDecrypter(undefined)
})

function settingsWith(overrides: {
  provider?: AiSettings['provider']
  anthropicKey?: string
  geminiKey?: string
  serperKey?: string
  tavilyKey?: string
  imageKey?: string
}): AiSettings {
  const settings = defaultAiSettings()
  settings.provider = overrides.provider ?? 'anthropic'
  if (overrides.anthropicKey !== undefined)
    settings.providers.anthropic!.apiKey = overrides.anthropicKey
  if (overrides.geminiKey !== undefined) settings.providers.gemini!.apiKey = overrides.geminiKey
  if (overrides.serperKey !== undefined)
    settings.search!.providers.serper.apiKey = overrides.serperKey
  if (overrides.tavilyKey !== undefined)
    settings.search!.providers.tavily.apiKey = overrides.tavilyKey
  if (overrides.imageKey !== undefined) {
    settings.media!.providers.openai.apiKey = overrides.imageKey
  }
  return settings
}

describe('settings secret masking', () => {
  it('masks long keys as first3 + ellipsis + last4', () => {
    expect(maskSecret('sk-ant-api03-abcdefgh')).toBe('sk-…efgh')
    expect(maskSecret('AIzaSyB1234567890')).toBe('AIz…7890')
  })

  it('uses the bullet placeholder for short keys and keeps empty empty', () => {
    expect(maskSecret('')).toBe('')
    expect(maskSecret('abc')).toBe(MASKED_SECRET_PLACEHOLDER)
    expect(maskSecret('1234567')).toBe(MASKED_SECRET_PLACEHOLDER)
    expect(maskSecret('12345678')).toBe('123…5678')
  })

  it('detects exactly the masked shapes, not real keys', () => {
    expect(isMaskedSecret(maskSecret('sk-ant-api03-abcdefgh'))).toBe(true)
    expect(isMaskedSecret(maskSecret('abc'))).toBe(true)
    expect(isMaskedSecret('')).toBe(false)
    expect(isMaskedSecret('sk-ant-api03-abcdefgh')).toBe(false)
    expect(isMaskedSecret('12345678')).toBe(false)
  })

  it('masks every secret section before settings reach a renderer', () => {
    const masked = maskAiSettingsSecrets(
      settingsWith({
        anthropicKey: 'sk-ant-api03-abcdefgh',
        serperKey: 'serper-secret-key-123',
        imageKey: 'sk-image-key-abcdefg',
      }),
    )
    expect(masked.providers.anthropic!.apiKey).toBe('sk-…efgh')
    expect(masked.search!.providers.serper.apiKey).toBe('ser…-123')
    expect(masked.media!.providers.openai.apiKey).toBe('sk-…defg')
    // non-secret fields survive untouched
    expect(masked.providers.anthropic!.model).toBe(defaultAiSettings().providers.anthropic!.model)
  })
})

describe('settings secret overlay', () => {
  it('keeps the stored key when the incoming value is masked or empty', () => {
    expect(overlaySecret('sk-…efgh', 'sk-ant-api03-abcdefgh')).toBe('sk-ant-api03-abcdefgh')
    expect(overlaySecret('', 'sk-ant-api03-abcdefgh')).toBe('sk-ant-api03-abcdefgh')
    expect(overlaySecret(MASKED_SECRET_PLACEHOLDER, 'short')).toBe('short')
  })

  it('replaces the key only for a real new value, and keeps empty when nothing is stored', () => {
    expect(overlaySecret('sk-brand-new-key', 'sk-ant-api03-abcdefgh')).toBe('sk-brand-new-key')
    expect(overlaySecret('', '')).toBe('')
    // nothing was stored, so nothing was ever masked: a mask-shaped value
    // typed into the empty field is a real (if odd) new key, not an echo
    expect(overlaySecret('sk-…efgh', '')).toBe('sk-…efgh')
  })

  it("treats a mask-shaped value as an echo only when it is the stored key's exact mask", () => {
    const stored = 'sk-ant-api03-abcdefgh'
    expect(overlaySecret('sk-…efgh', stored)).toBe(stored)
    // same 8-char ellipsis shape, different characters: a real new key
    expect(overlaySecret('AIz…7890', stored)).toBe('AIz…7890')
    // the placeholder still means unchanged for short stored keys
    expect(overlaySecret(MASKED_SECRET_PLACEHOLDER, 'short')).toBe('short')
    // a mask-shaped 8-char key can now be stored and replaced in turn
    // (maskSecret of 'xy…zwxy' is 'xy……zwxy', so 'sk-…efgh' is NOT its echo)
    const oddKey = 'xy…zwxy'
    expect(overlaySecret(oddKey, stored)).toBe(oddKey)
    expect(overlaySecret('sk-…efgh', oddKey)).toBe('sk-…efgh')
  })

  it('round-trips renderer edits over stored settings without losing other fields', () => {
    const stored = settingsWith({
      anthropicKey: 'sk-ant-api03-abcdefgh',
      geminiKey: 'AIzaSyB1234567890',
    })
    const incoming = settingsWith({ anthropicKey: 'sk-…efgh', geminiKey: 'AIz…7890' })
    incoming.providers.anthropic!.model = 'claude-opus-5'
    const merged = overlayAiSettingsSecrets(incoming, stored)
    expect(merged.providers.anthropic!.apiKey).toBe('sk-ant-api03-abcdefgh')
    expect(merged.providers.anthropic!.model).toBe('claude-opus-5')
    expect(merged.providers.gemini!.apiKey).toBe('AIzaSyB1234567890')
  })

  it('stores a brand-new key typed in the settings UI', () => {
    const stored = settingsWith({ anthropicKey: 'sk-ant-api03-abcdefgh' })
    const incoming = settingsWith({ anthropicKey: 'sk-totally-different' })
    expect(overlayAiSettingsSecrets(incoming, stored).providers.anthropic!.apiKey).toBe(
      'sk-totally-different',
    )
  })

  it('overlays media and search keys as well', () => {
    const stored = settingsWith({
      serperKey: 'serper-secret-key-123',
      imageKey: 'sk-image-key-abcdefg',
    })
    const incoming = settingsWith({ serperKey: 'ser…-123', imageKey: '' })
    const merged = overlayAiSettingsSecrets(incoming, stored)
    expect(merged.search!.providers.serper.apiKey).toBe('serper-secret-key-123')
    expect(merged.media!.providers.openai.apiKey).toBe('sk-image-key-abcdefg')
  })
})

describe('settings secret encryption at rest', () => {
  it('encrypts every secret into the enc: v1 envelope and decrypts it back', () => {
    const plain = settingsWith({
      anthropicKey: 'sk-ant-api03-abcdefgh',
      serperKey: 'serper-secret-key-123',
    })
    const encrypter = (value: string) => Buffer.from(value, 'utf-8').toString('base64')
    const stored = encryptStoredAiSettings(plain, encrypter)
    const storedProviders = stored.providers as Record<string, { apiKey: string }>
    expect(stored.v).toBe(1)
    expect(storedProviders.anthropic!.apiKey).toBe(
      `${ENCRYPTED_SECRET_PREFIX}c2stYW50LWFwaTAzLWFiY2RlZmdo`,
    )
    expect(isEncryptedSecret(storedProviders.anthropic!.apiKey)).toBe(true)

    setAiSecretDecrypter((blob) => Buffer.from(blob, 'base64').toString('utf-8'))
    const decrypted = decryptStoredAiSettings(stored)
    expect(decrypted.providers?.anthropic?.apiKey).toBe('sk-ant-api03-abcdefgh')
    expect(decrypted.search?.providers.serper.apiKey).toBe('serper-secret-key-123')
  })

  it('leaves empty keys unencrypted so defaults stay readable', () => {
    const stored = encryptStoredAiSettings(defaultAiSettings(), () => 'X')
    const storedProviders = stored.providers as Record<string, { apiKey: string }>
    expect(storedProviders.none!.apiKey).toBe('')
  })

  it('reads legacy plaintext files untouched and legacy enc: values through the decrypter', () => {
    const legacy = decryptStoredAiSettings({
      provider: 'custom',
      providers: {
        custom: { apiKey: 'plain-key', model: 'm', baseUrl: 'https://x/v1' },
      },
    })
    expect(legacy.providers?.custom?.apiKey).toBe('plain-key')

    setAiSecretDecrypter((blob) => Buffer.from(blob, 'base64').toString('utf-8'))
    const encryptedLegacy = decryptStoredAiSettings({
      apiKey: `${ENCRYPTED_SECRET_PREFIX}${Buffer.from('legacy-key').toString('base64')}`,
    })
    expect(encryptedLegacy.apiKey).toBe('legacy-key')
  })

  it('treats an undecryptable or decrypter-less enc: value as an unavailable key', () => {
    setAiSecretDecrypter(() => {
      throw new Error('OS keychain unavailable')
    })
    const failed = decryptStoredAiSettings({
      providers: { anthropic: { apiKey: `${ENCRYPTED_SECRET_PREFIX}not-really`, model: 'm' } },
    })
    expect(failed.providers?.anthropic?.apiKey).toBe('')

    const noDecrypter = decryptStoredAiSettings({
      providers: { anthropic: { apiKey: `${ENCRYPTED_SECRET_PREFIX}not-really`, model: 'm' } },
    })
    expect(noDecrypter.providers?.anthropic?.apiKey).toBe('')
  })
})
