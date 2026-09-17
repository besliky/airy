import { describe, expect, it } from 'vitest'

import {
  bridgeEnvDisabled,
  effectiveLiveBridgeEnabled,
  liveBridgeToggleAllowed,
} from '../src/main/live-bridge-state'

const ENV_ON = { AIRY_DISABLE_BRIDGE: '1' }
const ENV_OFF = { AIRY_DISABLE_BRIDGE: '' }

describe('bridgeEnvDisabled', () => {
  it('is set only by the exact sentinel value', () => {
    expect(bridgeEnvDisabled(ENV_ON)).toBe(true)
    expect(bridgeEnvDisabled(ENV_OFF)).toBe(false)
    expect(bridgeEnvDisabled({})).toBe(false)
    expect(bridgeEnvDisabled({ AIRY_DISABLE_BRIDGE: '0' })).toBe(false)
  })
})

describe('effectiveLiveBridgeEnabled', () => {
  it('treats an absent stored preference as enabled', () => {
    expect(effectiveLiveBridgeEnabled(undefined, ENV_OFF)).toBe(true)
    expect(effectiveLiveBridgeEnabled({}, ENV_OFF)).toBe(true)
  })

  it('follows the stored preference when the env override is absent', () => {
    expect(effectiveLiveBridgeEnabled(false, ENV_OFF)).toBe(false)
    expect(effectiveLiveBridgeEnabled(true, ENV_OFF)).toBe(true)
  })

  it('lets AIRY_DISABLE_BRIDGE=1 win over any stored preference', () => {
    expect(effectiveLiveBridgeEnabled(true, ENV_ON)).toBe(false)
    expect(effectiveLiveBridgeEnabled(undefined, ENV_ON)).toBe(false)
  })
})

describe('liveBridgeToggleAllowed', () => {
  it('allows toggling only when the env override is inactive', () => {
    expect(liveBridgeToggleAllowed(ENV_OFF)).toBe(true)
    expect(liveBridgeToggleAllowed({})).toBe(true)
    expect(liveBridgeToggleAllowed(ENV_ON)).toBe(false)
  })
})
