// Save channel for headless xlsx editing: the sheets app's streaming save
// gateway, reused as-is.
//
// Decision (S6): apps/sheets/src/gateway is verified Electron-free — its only
// non-local runtime dependencies are node builtins, jszip, zod, and the
// node-only @airy-office/electron-utils/atomic-write subpath (SA4
// section 5; the @airy-office/ai-provider / @airy-office/ui imports in
// shared/desktop-api are type-only and stripped at build time). Importing the
// app modules directly from this package therefore bundles cleanly (esbuild
// inlines them into dist/index.js; vitest resolves the raw TS through the
// workspace) and avoids forking ~13k lines of battle-tested byte-preserving
// patch logic — the CI preservation gate keeps protecting exactly the code we
// call. The inversion of the workspace direction (package importing app code)
// is deliberate and confined to this one module; if the app gateway ever
// gains an Electron import, this is the single place to swap in a vendored
// copy.
export {
  saveWorkbookViaSidecar,
  SaveTargetExistsError,
} from '../../../../apps/sheets/src/gateway/xlsx-package-io.js'
export type { CellEdit } from '../../../../apps/sheets/src/gateway/xlsx-gateway.js'
// Input schemas + types for agent-issued cell edits, straight from the save
// gateway's shared contract so the tool-accepted shape cannot drift from what
// the planner serializes (desktop-api is already part of the runtime bundle
// through xlsx-package-io's MAX_PATCH_ENTRY_BYTES import).
export type {
  WorkbookRichRun,
  WorkbookStyleEdit,
} from '../../../../apps/sheets/src/shared/desktop-api.js'
export {
  richRunSchema,
  workbookStyleEditSchema,
} from '../../../../apps/sheets/src/shared/desktop-api.js'
