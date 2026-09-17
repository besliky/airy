/**
 * Encrypted AI settings store, shared by every editor main process (docs
 * registers the ai:* IPC for the whole shell; slides' and sheets' standalone
 * builds mirror the handlers and import this module).
 *
 * Secrets are encrypted at rest with Electron safeStorage when the OS offers
 * a keychain (DPAPI / Keychain / libsecret); the file is additionally chmod
 * 0600 so the plaintext fallback and the encrypted variant both stay
 * user-readable only — the same care the MCP bridge token file gets. The
 * file keeps the v1 envelope `{ v: 1, providers: { … apiKey: 'enc:<base64>' } }`;
 * a legacy plaintext file is read transparently and re-encrypted on the next
 * save.
 *
 * Renderers only ever see masked keys (sk-…abcd). A masked or empty value
 * sent back means "unchanged"; real values replace the stored key. Main-side
 * request paths (stream/chat/test) overlay the stored real keys before use.
 */
import { app, safeStorage } from 'electron'
import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import {
  activeProvider,
  decryptStoredAiSettings,
  defaultAiSettings,
  encryptStoredAiSettings,
  isMaskedSecret,
  maskAiSettingsSecrets,
  overlayAiSettingsSecrets,
  overlaySecret,
  resolveAiSettings,
  setAiSecretDecrypter,
  type AiMediaProviderConfig,
  type AiMediaProviderId,
  type AiSearchProviderId,
  type AiSettings,
  type SecretEncrypter,
} from '@airy-office/ai-provider'

export function aiSettingsPath(): string {
  return join(app.getPath('userData'), 'ai-settings.json')
}

let codecRegistered = false

/**
 * Wire the safeStorage decrypter into @airy-office/ai-provider so file readers
 * outside this module (ai-search tools in docs/slides/sheets/markdown/html/pdf)
 * decode encrypted keys too. Idempotent; called by every load/save.
 */
export function registerAiSettingsCodec(): void {
  if (codecRegistered) return
  codecRegistered = true
  setAiSecretDecrypter((encryptedBase64) => {
    const plain = safeStorage.decryptString(Buffer.from(encryptedBase64, 'base64'))
    if (!plain) throw new Error('safeStorage returned no plaintext')
    return plain
  })
}

function encryptionAvailable(): boolean {
  try {
    return safeStorage.isEncryptionAvailable()
  } catch {
    return false
  }
}

const identityEncrypter: SecretEncrypter = (plain) => plain

function safeStorageEncrypter(plain: string): string {
  return safeStorage.encryptString(plain).toString('base64')
}

function readRawStored(): unknown {
  try {
    const path = aiSettingsPath()
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8'))
  } catch {
    /* corrupted settings file: fall through to defaults */
  }
  return {}
}

/** Stored settings with real (decrypted) keys, defaults merged, provider resolved. */
export function loadStoredAiSettings(): AiSettings {
  registerAiSettingsCodec()
  const settings = resolveAiSettings(decryptStoredAiSettings(readRawStored()), defaultAiSettings())
  settings.provider = activeProvider(settings)
  return settings
}

/** Settings for renderers: every secret masked (sk-…abcd / ••••). */
export function maskedAiSettings(): AiSettings {
  return maskAiSettingsSecrets(loadStoredAiSettings())
}

function writeJson0600(path: string, value: unknown): void {
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
  try {
    chmodSync(path, 0o600)
  } catch {
    /* Windows and some filesystems ignore the mode; the write itself stands */
  }
}

/**
 * Persist renderer-supplied settings: masked/empty secrets keep their stored
 * value, real ones replace it. Encrypted with safeStorage when available;
 * otherwise plaintext (still 0600) — losing the OS keychain must not lock the
 * user out of their own settings file.
 */
export function saveAiSettings(incoming: AiSettings): void {
  registerAiSettingsCodec()
  const merged = overlayAiSettingsSecrets(incoming, loadStoredAiSettings())
  const encrypter = encryptionAvailable() ? safeStorageEncrypter : identityEncrypter
  writeJson0600(aiSettingsPath(), encryptStoredAiSettings(merged, encrypter))
}

/**
 * Restore real stored keys onto renderer-supplied settings before a
 * main-process request (stream/chat): masked values mean the renderer echoed
 * what ai:get-settings gave it, empty means the field was never filled.
 */
export function overlayRealAiSecrets(settings: AiSettings): AiSettings {
  return overlayAiSettingsSecrets(settings, loadStoredAiSettings())
}

/** Single media-provider key overlay for the ai:media-test channel. */
export function realMediaApiKey(
  provider: AiMediaProviderId,
  config: AiMediaProviderConfig,
): AiMediaProviderConfig {
  if (!isMaskedSecret(config.apiKey) && config.apiKey !== '') return config
  const stored = loadStoredAiSettings().media?.providers[provider]?.apiKey ?? ''
  return { ...config, apiKey: overlaySecret(config.apiKey, stored) }
}

/** Single search-provider key overlay for the ai:search-test channel. */
export function realSearchApiKey(provider: AiSearchProviderId, apiKey: string): string {
  if (!isMaskedSecret(apiKey) && apiKey !== '') return apiKey
  return overlaySecret(apiKey, loadStoredAiSettings().search?.providers[provider]?.apiKey ?? '')
}
