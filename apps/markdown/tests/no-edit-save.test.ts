/**
 * OBS-1746 (regression of #196): a Ctrl+S with zero edits rewrote the file
 * anyway — the parse→serialize round trip is not byte-faithful for every
 * envelope (a mid-file BOM is filtered on display, code-block indentation is
 * re-serialized), so a CR-only/BOM file another tool wrote came back mutated
 * from a "save" that changed nothing. The chosen contract: an in-place save
 * of a clean document does not write at all — the on-disk bytes stay exactly
 * as the last writer left them. Menu Save already short-circuits clean views
 * main-side (requestMarkdownSave); Save As stays explicit because a picked
 * target is a deliberate write.
 *
 * The guard lives in the renderer's doSave (only the renderer knows its dirty
 * flag), so the behavioral bytes-identical outcome follows from the guard
 * itself; these source pins keep the contract connected the same way the
 * other renderer wiring tests do (cf. encoding-ui-wiring.test.ts).
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const renderer = readFileSync(join(here, '../src/renderer/App.tsx'), 'utf8')
const main = readFileSync(join(here, '../src/main/markdown-main.ts'), 'utf8')

describe('markdown no-edit save is byte-faithful (OBS-1746)', () => {
  it('a clean in-place save stops before any write', () => {
    // the exact guard: mode 'save', document not dirty, document has a path
    expect(renderer).toContain(
      "if (mode === 'save' && !dirtyRef.current && filePathRef.current) return true",
    )
    // …and it sits before serialization/IPC would run (the doSave body)
    const doSaveStart = renderer.indexOf('const doSave = useCallback(')
    expect(doSaveStart).toBeGreaterThan(0)
    const guardAt = renderer.indexOf(
      "if (mode === 'save' && !dirtyRef.current && filePathRef.current) return true",
    )
    const serializeAt = renderer.indexOf('serializeDocText(envelopeRef.current', doSaveStart)
    expect(guardAt).toBeGreaterThan(doSaveStart)
    expect(guardAt).toBeLessThan(serializeAt)
  })

  it('the guard is scoped to in-place saves of pathed documents', () => {
    // Save As and the untitled first save stay explicit: a picked target is a
    // deliberate write, and an untitled document has no bytes to preserve
    expect(renderer).toContain("if (mode === 'save' && !dirtyRef.current && filePathRef.current)")
  })

  it('the menu-save clean short-circuit main-side still backs the contract', () => {
    expect(main).toContain(
      "if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id))",
    )
  })
})
