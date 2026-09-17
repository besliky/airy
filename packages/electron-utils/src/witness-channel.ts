/// Channel the preload-world drop/paste witness reports resolved paths on.
///
/// This constant lives in its own dependency-free module on purpose:
/// `drop-open.ts` runs inside sandboxed preloads, where Node built-ins
/// (`node:path`, `node:fs`, …) cannot be required — any transitive import
/// from it into a module that uses them breaks the whole preload at load
/// ("module not found: node:path") and leaves the renderer without its
/// contextBridge API. Keep this file import-free.
export const WITNESS_DROP_CHANNEL = 'app:witnessed-dropped-files'
