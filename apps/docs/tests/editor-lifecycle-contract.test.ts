import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import {
  createTrackedEditor,
  drainTrackedEditors,
  trackedEditorCount,
} from './helpers/tracked-editor'

/**
 * TEST-1201 editor-lifecycle contract: every editor a docs test mounts must
 * be destroyed before its jsdom environment goes away. A leaked editor keeps
 * a ProseMirror DOMObserver polling timer alive across environment teardown,
 * where it throws an unhandled "document is not defined" that fails the WHOLE
 * run — the incident class that turned main red after green PRs twice
 * (PR #77, hotfix #80). This source-level test is the anti-recurrence gate.
 *
 * Mechanical traits, per `new Editor(` / `new EditorView(` site (modeled on
 * modal-dialog-contract.test.ts). A site is disciplined if ANY holds:
 *
 *   1. WRAPPED — the constructor is an argument of a tracking call, spelled
 *      on the line above (`track(`) or inline (`editors.add(new Editor(`) —
 *      the ai-rewrite-formatting idiom;
 *   2. REGISTERED — the assigned variable is pushed into a collection within
 *      24 lines below the site (`openEditors.push(editor)`,
 *      `editors.add(reopened)`, `track(editor)`), the hotfix #80 idiom;
 *   3. VAR-DESTROYED — the same variable is destroyed somewhere in the file
 *      (`editor.destroy()`), the per-instance revisions.test.ts idiom;
 *   4. BARE-RETURN — the site has no assigned variable (a helper's
 *      `return new Editor(...)`) and the file carries `.destroy()` at all.
 *
 * The preferred form since TEST-1202 is the shared helper
 * (tests/helpers/tracked-editor.ts): a `createTrackedEditor(` site hides the
 * constructor inside the helper, so it is disciplined iff the file drains
 * the registry (`drainTrackedEditors(` in an afterEach).
 *
 * Plus the file-level rule: a file that mounts editors must contain
 * `.destroy()` or call `drainTrackedEditors(` — a tracking collection nobody
 * drains (the cross-ref-modal regression) is a leak even though every site
 * looks registered.
 *
 * A new legitimate idiom means teaching this list, not bypassing the gate.
 */
const TESTS = join(__dirname, '.')

/** how far below a site the registration may sit (long `content:` literals) */
const REGISTRATION_WINDOW = 24

const listTestFiles = (dir: string): string[] => {
  const out: string[] = []
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name)
    if (entry.isDirectory()) out.push(...listTestFiles(path))
    else if (entry.name.endsWith('.test.ts')) out.push(path)
  }
  return out
}

interface Discipline {
  violations: string[]
}

/** every `new Editor(`/`new EditorView(`/`createTrackedEditor(` site that
 *  fails the traits above */
const scan = (): Discipline => {
  const violations: string[] = []
  for (const file of listTestFiles(TESTS)) {
    const src = readFileSync(file, 'utf8')
    const lines = src.split('\n')
    const hasDestroy = /\.destroy\(\)/.test(src) || /drainTrackedEditors\(/.test(src)
    // helper sites: disciplined exactly when the registry is drained here
    const helperSites: number[] = []
    lines.forEach((line, i) => {
      if (/createTrackedEditor\(/.test(line)) helperSites.push(i)
    })
    const sites: Array<{ i: number; kind: string }> = []
    lines.forEach((line, i) => {
      for (const m of line.matchAll(/new (Editor|EditorView)\(/g)) sites.push({ i, kind: m[1] })
    })
    if (!helperSites.length && !sites.length) continue
    const rel = file.slice(file.indexOf('tests/'))
    if (!hasDestroy)
      violations.push(
        `${rel}: mounts ${helperSites.length + sites.length} editor(s) but neither destroys them nor drains the tracked registry`,
      )
    if (!/drainTrackedEditors\(/.test(src))
      for (const i of helperSites)
        violations.push(
          `${rel}:${i + 1}: createTrackedEditor( registers an editor nobody drains — call drainTrackedEditors() in an afterEach`,
        )
    for (const { i, kind } of sites) {
      const line = lines[i]
      const prefix = line.slice(0, line.indexOf(`new ${kind}(`))
      const assigned = /(?:const|let|var)\s+(\w+)\s*=\s*$/.exec(prefix.trim())
      const variable = assigned ? assigned[1] : null
      let disciplined = false
      // 1. WRAPPED: tracking call opening on the line above or inline
      const prev = i > 0 ? lines[i - 1].trim() : ''
      if (/\btrack\(\s*$/.test(prev) || /\.(?:push|add)\(\s*$/.test(prefix)) disciplined = true
      // 2. REGISTERED: the variable joins a tracked collection below the site
      if (!disciplined && variable) {
        const registered = new RegExp(`\\.(?:push|add)\\(\\s*${variable}\\s*\\)`)
        const tracked = new RegExp(`\\btrack\\(\\s*${variable}\\s*[),]`)
        for (let j = i + 1; j <= Math.min(i + REGISTRATION_WINDOW, lines.length - 1); j++) {
          if (registered.test(lines[j]) || tracked.test(lines[j])) {
            disciplined = true
            break
          }
          // a new scope starting before any registration means it never comes
          if (/^\s*(?:describe|it)\(/.test(lines[j])) break
        }
      }
      // 3. VAR-DESTROYED: per-instance destroy of this variable in the file
      if (!disciplined && variable && new RegExp(`\\b${variable}\\.destroy\\(\\)`).test(src))
        disciplined = true
      // 4. BARE-RETURN: helper-return site, the file carries destroy discipline
      if (!disciplined && !variable && hasDestroy) disciplined = true
      if (!disciplined)
        violations.push(
          `${rel}:${i + 1}: new ${kind}( is never registered or destroyed — track it (openEditors.push) or destroy it`,
        )
    }
  }
  return { violations }
}

describe('editor lifecycle discipline (TEST-1201)', () => {
  it('finds the tests directory (scanner sanity)', () => {
    expect(listTestFiles(TESTS).length).toBeGreaterThan(100)
  })

  it('every mounted editor is tracked or destroyed', () => {
    const { violations } = scan()
    expect(violations).toEqual([])
  })
})

describe('tracked-editor helper registry (TEST-1202)', () => {
  it('registers created editors and drains them all', async () => {
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    expect(trackedEditorCount()).toBe(0)
    const editor = createTrackedEditor({ extensions: editorExtensions })
    expect(trackedEditorCount()).toBe(1)
    expect(editor.isDestroyed).toBe(false)
    await drainTrackedEditors()
    expect(trackedEditorCount()).toBe(0)
    expect(editor.isDestroyed).toBe(true)
  })
})
