// Bundles the MCP server into a single self-contained dist/index.js that runs
// with plain `node` and needs no node_modules (repo packages ship raw TS, so
// nothing here is runnable by tsc alone).
//
// ESM format: matches the package "type": "module", keeps import.meta.url
// working for the main-module check, and is safe because every dependency is
// inlined — there are no externals except node builtins, which esbuild
// externalizes automatically for platform=node.
//
// The banner defines a real CJS `require` for the inlined CommonJS
// dependencies (word-extractor via @airy-office/file-parse calls
// require('buffer') at load time); without it esbuild's ESM interop shim
// throws "Dynamic require of X is not supported".
import { build } from 'esbuild'

await build({
  entryPoints: ['src/index.ts'],
  outfile: 'dist/index.js',
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'esm',
  sourcemap: true,
  logLevel: 'info',
  banner: {
    js: "import { createRequire } from 'node:module'\nconst require = createRequire(import.meta.url)",
  },
})
