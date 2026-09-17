import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// Resolve workspace packages from this checkout's sources: in a git worktree
// node_modules is a symlink into the main checkout, so bare specifiers would
// silently bundle the other checkout's (possibly stale) code.
const localAlias = {
  '@airy-office/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
}

// Stable vendor chunks for the ~6 MB renderer bundle (docs shipped one
// monolithic index-*.js). Workspace packages resolve to their real
// packages/<name>/src paths (vite follows the node_modules symlinks), npm
// packages keep node_modules in the id — both are matched on the path, so it
// works on posix and win32 packagers alike.
const nodeModule = (name: string): RegExp =>
  new RegExp(`[\\\\/]node_modules[\\\\/]${name}([\\\\/]|$)`)
const workspacePkg = (name: string): RegExp => new RegExp(`[\\\\/]packages[\\\\/]${name}[\\\\/]`)
const rendererManualChunks = (id: string): string | undefined => {
  if (nodeModule('(react|react-dom|scheduler)').test(id)) return 'react'
  if (nodeModule('(@tiptap|prosemirror-[a-z-]+)').test(id)) return 'editor'
  if (id.includes('node_modules')) return 'vendor'
  if (workspacePkg('docx-engine').test(id)) return 'docx-engine'
  if (
    workspacePkg(
      '(ui|i18n|agent-core|ai-provider|ai-search|file-parse|project-store|pptx-render|electron-utils|font-metrics)',
    ).test(id)
  )
    return 'workspace'
  return undefined // app code stays in the entry chunk
}

export default defineConfig({
  // Main and preload use only electron + node builtins; bundle everything so
  // the packaged app doesn't rely on node_modules at runtime.
  // @airy-office/* deps ship as raw TS source with extensionless imports, so they
  // must be bundled — externalizing them yields ERR_MODULE_NOT_FOUND under Node
  // (same setup as apps/slides).
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@airy-office/electron-utils', '@airy-office/font-metrics'],
      }),
    ],
    resolve: { alias: localAlias },
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the drop-open bridge must be bundled, not externalized.
    plugins: [externalizeDepsPlugin({ exclude: ['@airy-office/electron-utils'] })],
  },
  renderer: {
    plugins: [react()],
    resolve: { alias: localAlias },
    build: {
      rollupOptions: {
        output: { manualChunks: rendererManualChunks },
      },
      // Documents current reality: the largest chunk is the entry (the app's
      // own editor code, ~3.1 MB); vendor groups all sit under ~0.8 MB. Note
      // vite's "chunks are larger than" warning never prints under
      // electron-vite (its renderer environment is not consumer 'client'), so
      // this limit is documentation until that changes.
      chunkSizeWarningLimit: 3200,
    },
    server: {
      // Overridable so multiple airy dev instances can coexist (default 5173).
      port: Number(process.env.DOCS_DEV_PORT) || 5173,
      strictPort: Boolean(process.env.DOCS_DEV_PORT),
    },
  },
})
