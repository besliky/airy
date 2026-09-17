import { shell } from 'electron'

/** the public open-source repository; also the target of every in-app star CTA */
export const GITHUB_REPO_URL = 'https://github.com/besliky/airy'

/** user-facing documentation: the repo README (there is no docs site yet) */
export const DOCS_README_URL = `${GITHUB_REPO_URL}/#readme`

/** the MCP copilot guide shipped in the repo (docs/COPILOT.md) */
export const COPILOT_GUIDE_URL = `${GITHUB_REPO_URL}/blob/main/docs/COPILOT.md`

/**
 * Open a help/docs link in the system browser through the suite's single
 * validation gate (safeExternalUrl): constants are used for these menus, but
 * every shell.openExternal must still pass the allowlist so a future edit that
 * routes user input here cannot reach file: or custom-scheme handlers.
 */
export async function openHelpUrl(url: string): Promise<void> {
  const { safeExternalUrl } = await import('./safe-external-url')
  const target = safeExternalUrl(url)
  if (target) await shell.openExternal(target).catch(() => {})
}
