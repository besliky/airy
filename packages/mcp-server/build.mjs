// Bundles the MCP server into a single self-contained dist/index.js that runs
// with plain `node` and needs no node_modules (repo packages ship raw TS, so
// nothing here is runnable by tsc alone).
//
// ESM format: matches the package "type": "module", keeps import.meta.url
// working for the main-module check, and is safe because every dependency is
// inlined — there are no externals except node builtins, which esbuild
// externalizes automatically for platform=node.
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
})
