import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
    // The smoke test spawns the bundled server and runs a stdio handshake
    testTimeout: 20000,
  },
})
