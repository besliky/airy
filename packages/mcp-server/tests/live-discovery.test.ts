// Unit tests for bridge info-file discovery: the AIRY_BRIDGE_FILE override
// (exclusive — no fallback) and the per-platform default search paths.
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import {
  BRIDGE_APP_NAME_CANDIDATES,
  BRIDGE_INFO_FILE_ENV,
  BRIDGE_INFO_NAME,
  candidateBridgeInfoPaths,
  discoverBridgeInfo,
} from '../src/live/discovery.js'

let home: string

beforeEach(async () => {
  home = await mkdtemp(join(tmpdir(), 'airy-discovery-'))
})

afterEach(async () => {
  await rm(home, { recursive: true, force: true })
})

const noOverrides = { XDG_CONFIG_HOME: '', APPDATA: '' }

/** write an info file into ~/.config/<appName> (the linux candidate layout) */
async function writeLinuxCandidate(appName: string, payload: string): Promise<string> {
  const path = join(home, '.config', appName, BRIDGE_INFO_NAME)
  await mkdir(join(home, '.config', appName), { recursive: true })
  await writeFile(path, payload, 'utf8')
  return path
}

function validInfo() {
  return { socketPath: '/run/airy-bridge.sock', token: 't'.repeat(64), pid: 1, protocolVersion: 1 }
}

describe('candidate info-file paths', () => {
  it('uses AIRY_BRIDGE_FILE exclusively when set', () => {
    const paths = candidateBridgeInfoPaths({
      env: { ...noOverrides, [BRIDGE_INFO_FILE_ENV]: '/custom/airy-bridge.json' },
      platform: 'linux',
      homeDir: home,
    })
    expect(paths).toEqual(['/custom/airy-bridge.json'])
  })

  it('ignores an empty AIRY_BRIDGE_FILE and falls back to defaults', () => {
    const paths = candidateBridgeInfoPaths({
      env: { ...noOverrides, [BRIDGE_INFO_FILE_ENV]: '  ' },
      platform: 'linux',
      homeDir: home,
    })
    expect(paths.length).toBe(BRIDGE_APP_NAME_CANDIDATES.length)
  })

  it('searches every app-name candidate under the platform appData base', () => {
    const linux = candidateBridgeInfoPaths({
      env: noOverrides,
      platform: 'linux',
      homeDir: '/home/u',
    })
    expect(linux).toEqual(
      BRIDGE_APP_NAME_CANDIDATES.map((name) => join('/home/u/.config', name, BRIDGE_INFO_NAME)),
    )

    const mac = candidateBridgeInfoPaths({
      env: noOverrides,
      platform: 'darwin',
      homeDir: '/Users/u',
    })
    expect(mac[0]).toBe(
      join('/Users/u/Library/Application Support', BRIDGE_APP_NAME_CANDIDATES[0], BRIDGE_INFO_NAME),
    )

    const win = candidateBridgeInfoPaths({
      env: { ...noOverrides, APPDATA: 'C:\\Users\\u\\AppData\\Roaming' },
      platform: 'win32',
      homeDir: 'C:\\Users\\u',
    })
    expect(win[0]).toBe(
      join('C:\\Users\\u\\AppData\\Roaming', BRIDGE_APP_NAME_CANDIDATES[0], BRIDGE_INFO_NAME),
    )
  })

  it('honors XDG_CONFIG_HOME on linux and APPDATA on windows', () => {
    expect(
      candidateBridgeInfoPaths({
        env: { ...noOverrides, XDG_CONFIG_HOME: '/xdg' },
        platform: 'linux',
        homeDir: home,
      })[0],
    ).toBe(join('/xdg', BRIDGE_APP_NAME_CANDIDATES[0], BRIDGE_INFO_NAME))

    expect(
      candidateBridgeInfoPaths({
        env: { XDG_CONFIG_HOME: '', APPDATA: 'D:\\Roaming' },
        platform: 'win32',
        homeDir: home,
      })[0],
    ).toBe(join('D:\\Roaming', BRIDGE_APP_NAME_CANDIDATES[0], BRIDGE_INFO_NAME))
  })
})

describe('discoverBridgeInfo', () => {
  it('returns null when no candidate exists', async () => {
    expect(
      await discoverBridgeInfo({ env: noOverrides, platform: 'linux', homeDir: home }),
    ).toBeNull()
  })

  it('returns null for a set-but-missing env override (no default fallback)', async () => {
    expect(
      await discoverBridgeInfo({
        env: { ...noOverrides, [BRIDGE_INFO_FILE_ENV]: join(home, 'missing.json') },
        platform: 'linux',
        homeDir: home,
      }),
    ).toBeNull()
  })

  it('picks the first existing candidate in app-name order', async () => {
    const [, second, third] = BRIDGE_APP_NAME_CANDIDATES
    const secondPath = await writeLinuxCandidate(second, JSON.stringify(validInfo()))
    await writeLinuxCandidate(third, JSON.stringify({ ...validInfo(), pid: 3 }))
    const found = await discoverBridgeInfo({ env: noOverrides, platform: 'linux', homeDir: home })
    expect(found?.path).toBe(secondPath)
    expect(found?.info).toEqual(validInfo())
  })

  it('skips malformed candidates and takes the next valid one', async () => {
    const [first, second] = BRIDGE_APP_NAME_CANDIDATES
    await writeLinuxCandidate(first, '{not json')
    const secondPath = await writeLinuxCandidate(second, JSON.stringify(validInfo()))
    const found = await discoverBridgeInfo({ env: noOverrides, platform: 'linux', homeDir: home })
    expect(found?.path).toBe(secondPath)
    expect(found?.info.socketPath).toBe('/run/airy-bridge.sock')
  })

  it('rejects structurally invalid payloads (missing token, wrong types)', async () => {
    const [first] = BRIDGE_APP_NAME_CANDIDATES
    const info = validInfo()
    for (const payload of [
      JSON.stringify({ ...info, token: '' }),
      JSON.stringify({ ...info, pid: 'one' }),
      JSON.stringify({ socketPath: '/s', token: 't', pid: 1 }), // missing protocolVersion
      JSON.stringify([info]),
      'null',
    ]) {
      await writeLinuxCandidate(first, payload)
      expect(
        await discoverBridgeInfo({ env: noOverrides, platform: 'linux', homeDir: home }),
      ).toBeNull()
      await rm(join(home, '.config'), { recursive: true, force: true })
    }
  })
})
