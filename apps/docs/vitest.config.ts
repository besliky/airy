import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

// resolve sibling source packages by path (not via node_modules), so a git
// worktree whose node_modules is linked to another checkout still tests
// against this checkout's edits (same convention as packages/pdf2docx)
const local = (rel: string) => fileURLToPath(new URL(rel, import.meta.url))

export default defineConfig({
  resolve: {
    alias: {
      '@airy-office/docx-engine': local('../../packages/docx-engine/src/index.ts'),
      '@airy-office/font-metrics': local('../../packages/font-metrics/src/index.ts'),
      '@airy-office/electron-utils': local('../../packages/electron-utils/src/index.ts'),
      '@airy-office/ai-provider/browser': local('../../packages/ai-provider/src/browser.ts'),
      '@airy-office/ai-provider': local('../../packages/ai-provider/src/index.ts'),
      '@airy-office/i18n': local('../../packages/i18n/src/index.ts'),
      '@airy-office/ui': local('../../packages/ui/src/index.ts'),
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
