import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

export default defineConfig({
  main: {
    // @airy-office/* workspace packages ship TS source (no build step, no
    // compiled entry point) — externalizing them makes Node's ESM loader try
    // to resolve their relative imports at runtime and fail. Bundle those;
    // externalize everything else (Electron, zod, node builtins).
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@airy-office/ai-provider',
          '@airy-office/agent-core',
          '@airy-office/ai-search',
          '@airy-office/docx-engine',
          '@airy-office/file-parse',
          '@airy-office/electron-utils',
          '@airy-office/i18n',
        ],
      }),
    ],
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the drop-open bridge must be bundled, not externalized.
    plugins: [externalizeDepsPlugin({ exclude: ['@airy-office/electron-utils'] })],
  },
  renderer: {
    plugins: [react()],
  },
})
