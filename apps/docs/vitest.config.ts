import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

// resolve sibling source packages by path (not via node_modules), so a git
// worktree whose node_modules is linked to another checkout still tests
// against this checkout's edits (same convention as packages/pdf2docx)
const local = (rel: string) => fileURLToPath(new URL(rel, import.meta.url))

export default defineConfig({
  resolve: {
    alias: [
      {
        find: '@airy-office/docx-engine',
        replacement: local('../../packages/docx-engine/src/index.ts'),
      },
      {
        find: '@airy-office/font-metrics',
        replacement: local('../../packages/font-metrics/src/index.ts'),
      },
      {
        find: '@airy-office/electron-utils',
        replacement: local('../../packages/electron-utils/src/index.ts'),
      },
      {
        find: '@airy-office/ai-provider/browser',
        replacement: local('../../packages/ai-provider/src/browser.ts'),
      },
      {
        find: '@airy-office/ai-provider',
        replacement: local('../../packages/ai-provider/src/index.ts'),
      },
      { find: '@airy-office/i18n', replacement: local('../../packages/i18n/src/index.ts') },
      // Exact match only: string aliases also rewrite subpaths, which broke
      // `@airy-office/ui/assets/airy-mark.png` (it became
      // `.../src/index.ts/assets/airy-mark.png`). Subpath imports fall through
      // to the workspace package's export map via the node_modules symlink,
      // the same way every other app resolves them.
      { find: /^@airy-office\/ui$/, replacement: local('../../packages/ui/src/index.ts') },
    ],
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
