import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

/// Vendor chunking for the sheets renderer (PERF-502).
///
/// The renderer used to build into a single 18.56 MB index-*.js because the
/// whole Univer preset stack is imported statically — every preset registers
/// its plugins inside `createUniver` before the first grid frame, so the
/// Univer graph cannot be deferred with dynamic imports without changing
/// plugin-registration timing. Splitting into the groups below only changes
/// file boundaries: the static import graph (and therefore module init
/// order) is preserved, the chunks are fetched in parallel via modulepreload.
///
/// The groups mirror Univer's own dependency layering (verified against the
/// `@univerjs/*` dist imports: core <- render/ui <- sheets <- features), so
/// the chunk graph stays acyclic — Rollup emits circular-chunk warnings (and
/// risks TDZ crashes at init) when grouping breaks that layering.
const LOCALE_DATA_RE = /\/locales?\// // dayjs `locale/`, Univer `locales/`
// engine-render ships ICU collation tables as lib/es/<lang>-<hash>.js files
// that only its locale loader fetches with dynamic import().
const ENGINE_RENDER_COLLATION_RE = /\/@univerjs\/engine-render\/lib\/es\/[^/]+-[^/]+\.js$/

/// First match wins; `univer-core` is the catch-all. Approximate weights:
const UNIVER_GROUPS: ReadonlyArray<readonly [name: string, match: RegExp]> = [
  // ~2 MB: sheet feature plugins (CF/DV/filter/sort/table/note/find-replace/
  // drawing) + their UI twins and preset wrappers. Everything here imports
  // the sheets/render groups below, so it must sit above them (e.g.
  // sheets-table-ui -> sheets-formula-ui, hence sheets-formula-ui lives
  // here and not in univer-sheets).
  [
    'univer-features',
    /\/@univerjs\/(?:preset-sheets-(?:conditional-formatting|data-validation|drawing|filter|find-replace|note|sort|table)|sheets-(?:conditional-formatting|data-validation|filter|formula-ui|sort|table|note|find-replace|drawing)(?:-ui)?|data-validation|find-replace)\//,
  ],
  // ~3.8 MB: sheet model + grid/ribbon UI + the sheets bridge of the formula
  // engine + number format. sheets and sheets-formula reference each other
  // (sheets -> engine-formula, sheets-formula -> sheets), so they share a
  // chunk; the quotient edge then points strictly at univer-formula.
  [
    'univer-sheets',
    /\/@univerjs\/(?:sheets-ui|sheets-formula(?!-ui)|sheets-numfmt(?:-ui)?|sheets)\//,
  ],
  // ~1.6 MB: the formula engine core (function registry + dependency graph)
  ['univer-formula', /\/@univerjs\/engine-formula\//],
  // ~1.9 MB: canvas render engine + shared component kit + the doc engine
  // that powers Univer's in-cell editor (docs/docs-ui) + drawing base
  [
    'univer-render-ui',
    /\/@univerjs\/(?:engine-render|ui|docs|docs-ui|docs-drawing|drawing|drawing-ui)\//,
  ],
  // ~1.2 MB: core, design tokens/themes, icons, rpc/protocol/network,
  // telemetry — the leaves every other Univer chunk imports
  ['univer-core', /\/@univerjs\//],
]

function rendererChunkOf(rawId: string): string | undefined {
  const id = rawId.replace(/\\/g, '/')
  if (LOCALE_DATA_RE.test(id) || ENGINE_RENDER_COLLATION_RE.test(id)) return undefined
  // PERF-901: the 19-locale app dictionary is not grouped here — i18n/strings.ts
  // dynamically imports i18n/locales/<lang>.ts, and leaving them ungrouped lets
  // Rollup emit one lazy chunk per locale (only the active one is fetched).
  if (!id.includes('/node_modules/')) return undefined
  // Preset wrappers sit at the very top of the graph (App.tsx imports them
  // and they reference every feature/sheets package): only the entry chunk
  // may import them, so they must not be pulled into a vendor group.
  if (/\/@univerjs\/preset-sheets-core\//.test(id)) return undefined
  if (/\/node_modules\/(?:react|react-dom|scheduler|react-is|use-sync-external-store)\//.test(id))
    return 'vendor-react'
  if (/\/node_modules\/rxjs\//.test(id)) return 'vendor-rxjs'
  for (const [name, match] of UNIVER_GROUPS) {
    if (match.test(id)) return name
  }
  // zod, jszip, opentype.js, unicode-regex, radix/fluentui, misc Univer deps
  return 'vendor-misc'
}

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
    build: {
      rollupOptions: {
        output: {
          manualChunks: rendererChunkOf,
        },
      },
    },
  },
})
