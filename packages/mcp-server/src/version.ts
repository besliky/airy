// Server identity shared by the entry point, the tool registry, and tests.
// The version is read from this package's package.json so the server always
// reports what was actually shipped (esbuild inlines the JSON into the bundle
// at build time; vitest resolves it through the workspace) — no second copy
// to fall out of sync with the package version.
import pkg from '../package.json'

export const SERVER_NAME = 'airy-mcp'

export const SERVER_VERSION: string = pkg.version
