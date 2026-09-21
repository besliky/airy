import { Editor, type EditorOptions } from '@tiptap/core'

/**
 * TEST-1202 shared tracked-editor helper — the one mounting idiom for docs
 * tests. An editor a test mounts and nobody destroys keeps a ProseMirror
 * DOMObserver polling timer alive across jsdom environment teardown, where
 * it throws an unhandled "document is not defined" that fails the WHOLE
 * vitest run (the incident class behind PR #77, hotfix #80 and TEST-1201;
 * it previously required a hand-rolled openEditors/liveEditors/track
 * collection per file).
 *
 * The idiom, in full:
 *
 *   import { createTrackedEditor, drainTrackedEditors } from './helpers/tracked-editor'
 *
 *   const editor = createTrackedEditor({ extensions, content }) // auto-tracked
 *   afterEach(() => drainTrackedEditors()) // destroy-all + settle window
 *
 * The registry is module-scoped and vitest isolates test files, so each
 * file's registry is its own. drainTrackedEditors destroys in reverse
 * mounting order, then waits 30ms for a pending DOMObserver flush to land
 * while the document still exists. editor-lifecycle-contract.test.ts
 * accepts this helper as a disciplined form: a file that mounts through
 * createTrackedEditor must call drainTrackedEditors.
 */

const tracked: Editor[] = []

/** `new Editor(...)` that registers itself for the afterEach drain. */
export function createTrackedEditor(options: Partial<EditorOptions> = {}): Editor {
  const { element = document.createElement('div'), ...rest } = options
  const editor = new Editor({ ...rest, element })
  tracked.push(editor)
  return editor
}

/** Destroy every tracked editor (reverse mounting order), then wait out a
 *  pending DOMObserver flush while the jsdom document still exists. */
export async function drainTrackedEditors(): Promise<void> {
  while (tracked.length) tracked.pop()?.destroy()
  await new Promise((resolve) => setTimeout(resolve, 30))
}

/** Registry size — the lifecycle contract's view into this module. */
export function trackedEditorCount(): number {
  return tracked.length
}
