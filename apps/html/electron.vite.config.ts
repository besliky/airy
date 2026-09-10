import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

export default defineConfig({
  // @airy-office/i18n and @airy-office/electron-utils ship as TS source — must be bundled
  main: {
    plugins: [
      externalizeDepsPlugin({ exclude: ['@airy-office/i18n', '@airy-office/electron-utils'] }),
    ],
  },
  preload: {
    plugins: [
      externalizeDepsPlugin({ exclude: ['@airy-office/i18n', '@airy-office/electron-utils'] }),
    ],
  },
  renderer: {
    plugins: [react()],
    server: {
      port: Number(process.env.HTML_DEV_PORT) || 5178,
      strictPort: Boolean(process.env.HTML_DEV_PORT),
    },
  },
})
