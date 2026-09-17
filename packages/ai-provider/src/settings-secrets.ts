/**
 * Secret handling for the persisted ai-settings.json: encryption at rest
 * (Electron safeStorage, wired in by the app main process via
 * setAiSecretDecrypter), masking for values handed to renderers, and the
 * overlay that restores real keys when a renderer echoes masked ones back.
 *
 * This module is pure TypeScript — no Electron import — so every rule it
 * encodes is unit-testable outside the app.
 */
import type { AiMediaSettings, AiSearchSettings, AiSettings, LegacyAiSettings } from './types'

/** Prefix marking an encrypted secret in the settings file (base64 blob after it). */
export const ENCRYPTED_SECRET_PREFIX = 'enc:'
/** Placeholder shown instead of secrets too short to mask without leaking them. */
export const MASKED_SECRET_PLACEHOLDER = '••••'
const ELLIPSIS = '…'

export type SecretDecrypter = (encryptedBase64: string) => string
export type SecretEncrypter = (plain: string) => string

let activeDecrypter: SecretDecrypter | undefined

/**
 * Register the safeStorage-backed decrypter (main process). Pure contexts
 * (tests, non-Electron consumers) run without one: encrypted values then read
 * as empty, i.e. "key not available".
 */
export function setAiSecretDecrypter(decrypter: SecretDecrypter | undefined): void {
  activeDecrypter = decrypter
}

export function isEncryptedSecret(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith(ENCRYPTED_SECRET_PREFIX)
}

/** sk-ant-api03-abcdefgh → "sk-…efgh"; empty stays empty; short keys become bullets. */
export function maskSecret(secret: string): string {
  const value = secret.trim()
  if (!value) return ''
  // Below 8 characters, first3+last4 would reveal the whole key.
  if (value.length < 8) return MASKED_SECRET_PLACEHOLDER
  return `${value.slice(0, 3)}${ELLIPSIS}${value.slice(-4)}`
}

/**
 * Detects the exact shape produced by maskSecret: the short placeholder, or
 * first3 + '…' + last4 (length 8, ellipsis at index 3). Shape alone cannot
 * prove the value is an echo — a real 8-character key containing '…' at
 * index 3 has the same shape — so the overlay additionally compares against
 * the stored key's actual mask (isMaskedSecretEcho) before treating a value
 * as "unchanged".
 */
export function isMaskedSecret(value: string): boolean {
  return (
    value === MASKED_SECRET_PLACEHOLDER ||
    (value.length === 8 && value.charCodeAt(3) === ELLIPSIS.charCodeAt(0))
  )
}

/** Whether `value` is the renderer echoing back the masked form of `stored`. */
function isMaskedSecretEcho(value: string, stored: string): boolean {
  if (value === MASKED_SECRET_PLACEHOLDER) return true
  if (!isMaskedSecret(value)) return false
  // the exact mask of the stored key; an empty stored key has no mask, so a
  // mask-shaped value typed into an empty field is a real new key
  return stored !== '' && value === maskSecret(stored)
}

/**
 * Single-secret overlay used by the save and connection-test paths: an empty
 * incoming value or an echo of the stored key's mask means "unchanged" (keep
 * the stored key); anything else is a real new key and replaces it.
 */
export function overlaySecret(incoming: string | undefined, stored: string | undefined): string {
  const value = incoming ?? ''
  const keep = stored ?? ''
  if (value === '' || isMaskedSecretEcho(value, keep)) return keep
  return value
}

function decryptValue(value: unknown, decrypter: SecretDecrypter | undefined): string | undefined {
  if (typeof value !== 'string') return undefined
  if (!isEncryptedSecret(value)) return value
  if (!decrypter) return ''
  try {
    return decrypter(value.slice(ENCRYPTED_SECRET_PREFIX.length))
  } catch {
    // Wrong OS key / moved file: the key is unavailable, not fatal.
    return ''
  }
}

function encryptValue(value: string, encrypter: SecretEncrypter): string {
  return `${ENCRYPTED_SECRET_PREFIX}${encrypter(value)}`
}

type Json = Record<string, unknown>

function asJson(value: unknown): Json | undefined {
  return typeof value === 'object' && value !== null ? (value as Json) : undefined
}

/** Decrypt the `apiKey` field of every entry in a providers record, in place. */
function decryptProviderSecrets(providers: unknown, decrypter: SecretDecrypter | undefined): void {
  const record = asJson(providers)
  if (!record) return
  for (const configValue of Object.values(record)) {
    const config = asJson(configValue)
    if (!config) continue
    const decrypted = decryptValue(config.apiKey, decrypter)
    if (decrypted !== undefined) config.apiKey = decrypted
  }
}

/**
 * Replace every `enc:` value in a stored settings object with plaintext using
 * the registered decrypter. Legacy plaintext values pass through untouched, so
 * one reader covers both file generations. The input is whatever JSON.parse
 * produced; the output is shaped for resolveAiSettings.
 */
export function decryptStoredAiSettings(stored: unknown): Partial<AiSettings> & LegacyAiSettings {
  const root = asJson(stored) ?? {}
  const decrypted: Json = { ...root }
  const providers = asJson(root.providers)
  if (providers) {
    const providersCopy: Json = { ...providers }
    decryptProviderSecrets(providersCopy, activeDecrypter)
    decrypted.providers = providersCopy
  }
  const media = asJson(root.media)
  if (media) {
    const mediaCopy: Json = { ...media }
    const mediaProviders = asJson(mediaCopy.providers)
    if (mediaProviders) {
      const mediaProvidersCopy: Json = { ...mediaProviders }
      decryptProviderSecrets(mediaProvidersCopy, activeDecrypter)
      mediaCopy.providers = mediaProvidersCopy
    }
    decrypted.media = mediaCopy
  }
  const search = asJson(root.search)
  if (search) {
    const searchCopy: Json = { ...search }
    const searchProviders = asJson(searchCopy.providers)
    if (searchProviders) {
      const searchProvidersCopy: Json = { ...searchProviders }
      decryptProviderSecrets(searchProvidersCopy, activeDecrypter)
      searchCopy.providers = searchProvidersCopy
    }
    decrypted.search = searchCopy
  }
  const legacyKey = decryptValue(root.apiKey, activeDecrypter)
  if (legacyKey !== undefined) decrypted.apiKey = legacyKey
  return decrypted as Partial<AiSettings> & LegacyAiSettings
}

function mapProviderKeys(
  providers: Record<string, { apiKey: string }>,
  map: (key: string) => string,
): Json {
  const out: Json = {}
  for (const [id, config] of Object.entries(providers)) {
    out[id] = { ...config, apiKey: map(config.apiKey) }
  }
  return out
}

/** Mask every secret before the settings cross into a renderer. */
export function maskAiSettingsSecrets(settings: AiSettings): AiSettings {
  const masked: Json = { ...settings }
  masked.providers = mapProviderKeys(settings.providers, maskSecret)
  if (settings.media) {
    masked.media = {
      ...settings.media,
      providers: mapProviderKeys(
        settings.media.providers as unknown as Record<string, { apiKey: string }>,
        maskSecret,
      ),
    }
  }
  if (settings.search) {
    masked.search = {
      ...settings.search,
      providers: mapProviderKeys(
        settings.search.providers as unknown as Record<string, { apiKey: string }>,
        maskSecret,
      ),
    }
  }
  return masked as unknown as AiSettings
}

/**
 * Merge incoming renderer settings over the stored real ones: masked or empty
 * secrets keep the stored key, real new values replace it. Used both when
 * saving and before main-process requests (stream/chat/test) so a renderer
 * that only ever saw masked keys still works.
 */
export function overlayAiSettingsSecrets(incoming: AiSettings, stored: AiSettings): AiSettings {
  const providers: Record<string, unknown> = {}
  const ids = new Set<string>([
    ...Object.keys(incoming.providers ?? {}),
    ...Object.keys(stored.providers ?? {}),
  ])
  for (const id of ids) {
    const from = incoming.providers?.[id as keyof typeof incoming.providers]
    const keep = stored.providers?.[id as keyof typeof stored.providers]
    if (from === undefined) {
      if (keep !== undefined) providers[id] = keep
      continue
    }
    if (keep === undefined) {
      providers[id] = from
      continue
    }
    providers[id] = { ...keep, ...from, apiKey: overlaySecret(from.apiKey, keep.apiKey) }
  }
  const overlayMedia = (): AiMediaSettings | undefined => {
    const media = incoming.media
    const base = stored.media
    if (!media) return base
    if (!base) return media
    const merged: Record<string, unknown> = {}
    const mediaIds = new Set<string>([
      ...Object.keys(media.providers ?? {}),
      ...Object.keys(base.providers ?? {}),
    ])
    for (const id of mediaIds) {
      const from = media.providers[id as keyof typeof media.providers]
      const keep = base.providers[id as keyof typeof base.providers]
      if (keep === undefined) merged[id] = from
      else if (from === undefined) merged[id] = keep
      else merged[id] = { ...keep, ...from, apiKey: overlaySecret(from.apiKey, keep.apiKey) }
    }
    return { ...base, ...media, providers: merged } as AiMediaSettings
  }
  const overlaySearch = (): AiSearchSettings | undefined => {
    const search = incoming.search
    const base = stored.search
    if (!search) return base
    if (!base) return search
    const merged: Record<string, unknown> = {}
    const searchIds = new Set<string>([
      ...Object.keys(search.providers ?? {}),
      ...Object.keys(base.providers ?? {}),
    ])
    for (const id of searchIds) {
      const from = search.providers[id as keyof typeof search.providers]
      const keep = base.providers[id as keyof typeof base.providers]
      merged[id] = { apiKey: overlaySecret(from?.apiKey, keep?.apiKey) }
    }
    return { ...base, ...search, providers: merged } as AiSearchSettings
  }
  const result: Json = { ...incoming, providers }
  if (incoming.media !== undefined || stored.media !== undefined) result.media = overlayMedia()
  if (incoming.search !== undefined || stored.search !== undefined) result.search = overlaySearch()
  return result as unknown as AiSettings
}

/**
 * Build the v1 stored envelope: same fields as AiSettings with every non-empty
 * secret replaced by `enc:` + the encrypter's blob. The Electron store decides
 * whether the encrypter is safeStorage or the identity (encryption
 * unavailable → plaintext, documented there). Returns the plain JSON shape
 * that gets written to disk.
 */
export function encryptStoredAiSettings(
  settings: AiSettings,
  encrypter: SecretEncrypter,
): Record<string, unknown> {
  const stored: Json = {
    v: 1,
    provider: settings.provider,
    providers: mapProviderKeys(settings.providers, (key) =>
      key ? encryptValue(key, encrypter) : key,
    ),
    ...(settings.maxOutputTokens !== undefined
      ? { maxOutputTokens: settings.maxOutputTokens }
      : {}),
  }
  if (settings.media) {
    stored.media = {
      ...settings.media,
      providers: mapProviderKeys(
        settings.media.providers as unknown as Record<string, { apiKey: string }>,
        (key) => (key ? encryptValue(key, encrypter) : key),
      ),
    }
  }
  if (settings.search) {
    stored.search = {
      ...settings.search,
      providers: mapProviderKeys(
        settings.search.providers as unknown as Record<string, { apiKey: string }>,
        (key) => (key ? encryptValue(key, encrypter) : key),
      ),
    }
  }
  return stored
}
