/** home:* channel access policy (src/main/home-channel-access.ts): every
 *  mutating / dialog- / tab-spawning home channel is Home-only; only pure
 *  reads stay open. Default-deny: an unclassified channel is Home-only. */
import { describe, expect, it } from 'vitest'

import { HOME_CHANNELS } from '../src/shared/home-api'
import {
  homeChannelAccess,
  homeHandlerAllowed,
  openHomeChannels,
} from '../src/main/home-channel-access'

const HOME_ID = 42
const EDITOR_ID = 77

describe('homeChannelAccess classification', () => {
  it('home-only: file mutations, settings writes, dialog/tab spawners', () => {
    const guarded = [
      HOME_CHANNELS.statPaths,
      HOME_CHANNELS.toggleStar,
      HOME_CHANNELS.openPath,
      HOME_CHANNELS.browse,
      HOME_CHANNELS.newDoc,
      HOME_CHANNELS.newSheet,
      HOME_CHANNELS.newSlide,
      HOME_CHANNELS.newMarkdown,
      HOME_CHANNELS.newHtml,
      HOME_CHANNELS.newPdf,
      HOME_CHANNELS.removeRecent,
      HOME_CHANNELS.revealPath,
      HOME_CHANNELS.renameFile,
      HOME_CHANNELS.duplicateFile,
      HOME_CHANNELS.deleteFiles,
      HOME_CHANNELS.openTrash,
      HOME_CHANNELS.setLanguage,
      HOME_CHANNELS.setOnboardingSeen,
      HOME_CHANNELS.setTheme,
      HOME_CHANNELS.setAuthorName,
      HOME_CHANNELS.setAutoSaveDefault,
      HOME_CHANNELS.setLiveBridgeEnabled,
      HOME_CHANNELS.setRestoreSession,
      HOME_CHANNELS.setAiPanelPrefs,
      // widens the renderer read allowlist (grantRendererDir): not an open read
      HOME_CHANNELS.getDefaultSaveDir,
      HOME_CHANNELS.pickDefaultSaveDir,
      HOME_CHANNELS.openGitHubRepo,
      // writes the shown/resolved state on grant
      HOME_CHANNELS.starPromptShouldShow,
      HOME_CHANNELS.starPromptAction,
    ]
    for (const channel of guarded) {
      expect(homeChannelAccess(channel), channel).toBe('home-only')
    }
  })

  it('open: the exempt read-only queries', () => {
    const open = [
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
    ]
    for (const channel of open) {
      expect(homeChannelAccess(channel), channel).toBe('open')
    }
  })

  it('every exemption names a real home:* channel (typos cannot hide)', () => {
    const real = new Set(Object.values(HOME_CHANNELS) as readonly string[])
    for (const channel of openHomeChannels()) {
      expect(real.has(channel), channel).toBe(true)
    }
    // and the exempt set is exactly the open classification
    expect(new Set(openHomeChannels()).size).toBe(openHomeChannels().length)
  })

  it('unclassified channels default to home-only (default-deny)', () => {
    expect(homeChannelAccess('home:some-future-channel')).toBe('home-only')
    expect(homeChannelAccess('')).toBe('home-only')
    expect(homeChannelAccess('app:get-theme')).toBe('home-only')
  })
})

describe('homeHandlerAllowed (fake senders)', () => {
  it('the Home tab webContents may call every home channel', () => {
    for (const channel of Object.values(HOME_CHANNELS) as readonly string[]) {
      expect(homeHandlerAllowed(channel, HOME_ID, HOME_ID), channel).toBe(true)
    }
  })

  it('an editor renderer is refused on home-only channels but keeps the open reads', () => {
    expect(homeHandlerAllowed(HOME_CHANNELS.deleteFiles, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.renameFile, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.setTheme, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.setLiveBridgeEnabled, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.newDoc, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.openTrash, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.pickDefaultSaveDir, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.starPromptAction, EDITOR_ID, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.recents, EDITOR_ID, HOME_ID)).toBe(true)
    expect(homeHandlerAllowed(HOME_CHANNELS.getAppVersion, EDITOR_ID, HOME_ID)).toBe(true)
    expect(homeHandlerAllowed(HOME_CHANNELS.getTheme, EDITOR_ID, HOME_ID)).toBe(true)
  })

  it('without a captured Home webContents even the Home renderer id is refused', () => {
    expect(homeHandlerAllowed(HOME_CHANNELS.browse, HOME_ID, null)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.deleteFiles, undefined, HOME_ID)).toBe(false)
    expect(homeHandlerAllowed(HOME_CHANNELS.recents, undefined, null)).toBe(true)
  })
})
