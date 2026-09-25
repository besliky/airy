import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    setupFiles: ['tests/setup-webstorage.ts', 'tests/setup-jsdom-geometry.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
