import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    setupFiles: ['tests/setup-webstorage.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
