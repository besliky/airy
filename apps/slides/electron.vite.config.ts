import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

// Pin resolution to this repo's workspace sources (matches tsconfig paths;
// avoids bundling stale implementations when node_modules links point elsewhere)
const workspaceAlias = {
  // Subpath before the bare name: string aliases are prefix replacements
  '@airy-office/pptx-engine/table-grid': resolve(
    here,
    '../../packages/pptx-engine/src/table-grid.ts',
  ),
  '@airy-office/pptx-engine/identity': resolve(here, '../../packages/pptx-engine/src/identity.ts'),
  '@airy-office/pptx-engine/custgeom': resolve(here, '../../packages/pptx-engine/src/custgeom.ts'),
  '@airy-office/pptx-engine/background-promote': resolve(
    here,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@airy-office/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
  '@airy-office/pptx-render/preset-geometry': resolve(
    here,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@airy-office/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
  // Metafile (EMF/WMF) rasterizer shared with the docs engine (renderer-only: needs canvas)
  '@airy-office/docx-engine/metafile': resolve(here, '../../packages/docx-engine/src/metafile.ts'),
}

// Stable vendor chunks for the ~5 MB renderer bundle (slides shipped one
// monolithic index-*.js). Workspace packages resolve to their real
// packages/<name>/src paths (vite follows the node_modules symlinks), npm
// packages keep node_modules in the id — both are matched on the path, so it
// works on posix and win32 packagers alike.
const nodeModule = (name: string): RegExp =>
  new RegExp(`[\\\\/]node_modules[\\\\/]${name}([\\\\/]|$)`)
const workspacePkg = (name: string): RegExp => new RegExp(`[\\\\/]packages[\\\\/]${name}[\\\\/]`)
const rendererManualChunks = (id: string): string | undefined => {
  if (nodeModule('(react|react-dom|scheduler)').test(id)) return 'react'
  if (nodeModule('(konva|react-konva)').test(id)) return 'konva'
  if (id.includes('node_modules')) return 'vendor'
  // the renderer pulls only slivers of the engines (parsing/saving live in
  // the main process), so they share one chunk instead of two tiny files
  if (workspacePkg('(pptx-engine|pptx-render)').test(id)) return 'pptx'
  if (
    workspacePkg(
      '(ui|i18n|agent-core|ai-provider|ai-search|file-parse|project-store|docx-engine|electron-utils)',
    ).test(id)
  )
    return 'workspace'
  return undefined // app code stays in the entry chunk
}

export default defineConfig({
  // Main process/preload must bundle @airy-office/* sources (they are pulled in as TS
  // source with extensionless relative imports; externalizing them under Node
  // yields ERR_MODULE_NOT_FOUND).
  main: {
    resolve: { alias: workspaceAlias },
    // Bundle opentype.js too (the packaged app ships only out/**, so external deps are unresolvable at runtime)
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@airy-office/pptx-engine',
          '@airy-office/pptx-render',
          '@airy-office/ai-search',
          '@airy-office/file-parse',
          '@airy-office/electron-utils',
          'opentype.js',
        ],
      }),
    ],
  },
  preload: {
    // electron-utils ships raw TS source — must be bundled, not left external
    plugins: [externalizeDepsPlugin({ exclude: ['@airy-office/electron-utils'] })],
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    build: {
      rollupOptions: {
        output: { manualChunks: rendererManualChunks },
      },
      // Documents current reality: the largest chunk is the entry (the app's
      // own editor code, ~3.0 MB); vendor groups all sit under ~0.6 MB. Note
      // vite's "chunks are larger than" warning never prints under
      // electron-vite (its renderer environment is not consumer 'client'), so
      // this limit is documentation until that changes.
      chunkSizeWarningLimit: 3000,
    },
    server: {
      port: Number(process.env.SLIDES_DEV_PORT) || 5175,
      strictPort: Boolean(process.env.SLIDES_DEV_PORT),
    },
  },
})
