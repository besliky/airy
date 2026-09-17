import { describe, expect, it } from 'vitest'
import { effectiveAuthorName, type ReviewContext } from '../src/renderer/review-actions'
import { setModuleLang } from '../src/renderer/i18n/locale'

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

// ---- effectiveAuthorName (the docs author-selection fallback) ----

describe('effectiveAuthorName', () => {
  it('prefers the configured author name', () => {
    expect(effectiveAuthorName('Ada Lovelace', 'User')).toBe('Ada Lovelace')
  })

  it('falls back to the localized default when unset or blank', () => {
    expect(effectiveAuthorName('', 'User')).toBe('User')
    expect(effectiveAuthorName('   ', 'User')).toBe('User')
  })

  it('trims the configured name before use', () => {
    expect(effectiveAuthorName('  Ada  ', 'User')).toBe('Ada')
  })

  it('an unset ctx.authorName behaves like the empty string', () => {
    const ctx = {} as Pick<ReviewContext, 'authorName'>
    expect(effectiveAuthorName(ctx.authorName ?? '', 'User')).toBe('User')
  })
})
