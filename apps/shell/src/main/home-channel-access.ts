/**
 * Access policy for the process-global home:* IPC channels.
 *
 * The shell bundles every editor's main code into one process, so an
 * ipcMain.handle registered for the Home screen answers ANY webContents — a
 * compromised editor renderer could otherwise rename/trash files, flip
 * settings, or spawn dialogs and tabs through them. Handlers are Home-only
 * BY DEFAULT (default-deny): only the channels explicitly listed here as
 * 'open' — pure reads of settings/lists with no side effects — skip the
 * sender check. The classification is data so tests can pin it: every
 * HOME_CHANNELS value must appear exactly once, and any channel a future
 * change forgets to classify stays Home-only.
 *
 * The only legitimate sender for guarded channels is the Home tab, which is
 * the shell window's own renderer (its webContents id is captured at window
 * creation). Editor renderers never invoke home:* — they read shared
 * settings through the separate app:get-* channels.
 */
import { HOME_CHANNELS } from '../shared/home-api'

import { isHomeSender } from './home-sender-guard'

export type HomeChannelAccess = 'home-only' | 'open'

/**
 * Channels exempt from the Home-sender check: read-only queries that neither
 * mutate persisted state, nor widen renderer file access, nor spawn dialogs,
 * windows, tabs, or external apps. Everything else is Home-only.
 */
const OPEN_HOME_CHANNELS: ReadonlySet<string> = new Set([
  HOME_CHANNELS.getAppVersion,
  HOME_CHANNELS.recents,
  HOME_CHANNELS.starred,
  HOME_CHANNELS.getLanguage,
  HOME_CHANNELS.onboardingSeen,
  HOME_CHANNELS.getTheme,
  HOME_CHANNELS.getAutoSaveDefault,
  HOME_CHANNELS.getAuthorName,
  HOME_CHANNELS.getLiveBridgeEnabled,
  HOME_CHANNELS.getLiveBridgeEnvDisabled,
  HOME_CHANNELS.getRestoreSession,
  HOME_CHANNELS.getAiPanelPrefs,
  HOME_CHANNELS.githubStars,
])

/** 'open' for the exempt reads, 'home-only' for every other home:* channel
 *  (including channels added after this table — default-deny). */
export function homeChannelAccess(channel: string): HomeChannelAccess {
  return OPEN_HOME_CHANNELS.has(channel) ? 'open' : 'home-only'
}

/** Guarded channels must come from the Home tab; open reads answer anyone.
 *  Pure so the decision is unit-testable with fake sender ids. */
export function homeHandlerAllowed(
  channel: string,
  senderId: number | undefined,
  homeWebContentsId: number | null | undefined,
): boolean {
  return homeChannelAccess(channel) === 'open' || isHomeSender(homeWebContentsId, senderId)
}

/** Test helper: the raw exemption list (tests pin that every entry names a
 *  real home:* channel — a typo would silently stop exempting it). */
export function openHomeChannels(): readonly string[] {
  return [...OPEN_HOME_CHANNELS]
}
