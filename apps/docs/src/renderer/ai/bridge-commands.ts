// Live-bridge command handlers (Airy Copilot): executes bridge requests from
// the shell's local socket server against THIS tab's editor by reusing the
// embedded agent pipeline — buildDocContext for context+selection, executeTool
// for insert_content/apply_ops (one ProseMirror transaction each, same
// stale-guard and tracked-changes behavior), and the AI panel's snapshot
// rollback for undo. One bridge call = one "turn"; undo reverts exactly the
// last turn and refuses when the user has edited since.
import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID, type Block } from '@genoffice/docx-engine'

import type { BridgeCommandResult } from '../../shared/ipc'
import { TABLE_TRAILING_SKIP } from '../editor/extensions'
import { buildDocContext, findNumId, getSelectionScope, type NumIds } from './protocol'
import {
  executeTool,
  markDocSeen,
  STALE_DOC_ERROR,
  type AiCommentsAccess,
  type AiHeaderFooterAccess,
  type ToolExecution,
} from './tools'

/** tracked-changes author for bridge edits (mirrors the AI panel's AI_REVISION_AUTHOR convention) */
export const BRIDGE_REVISION_AUTHOR = 'Airy Copilot'

/** must stay in sync with AiPanel's toggle (localStorage persistence) */
const TRACK_CHANGES_KEY = 'ai-docs-track-changes'

/** the document facts the bridge needs from the owning App */
export interface BridgeDocState {
  blocks: Block[]
  isBlank: boolean
  filePath: string | null
}

export interface BridgeCommandDeps {
  getEditor(): Editor | null
  /** null when this tab has no document open (bridge answers no_active_document) */
  getDocState(): BridgeDocState | null
  getComments?(): AiCommentsAccess | undefined
  getHf?(): AiHeaderFooterAccess | undefined
  /** track-changes toggle; defaults to the AI panel's persisted localStorage flag */
  isTrackChangesOn?(): boolean
}

/** the undo target: doc state right around the last bridge-initiated mutation */
interface BridgeTurn {
  /** editor JSON captured before the turn's mutation */
  before: JSONContent
  /** the doc node right after the mutation — undo only runs while it is still current */
  afterDoc: PmNode
}

const fail = (code: string, message: string): BridgeCommandResult => ({
  ok: false,
  error: { code, message },
})

/** the AI panel's persisted toggle; storage can be missing in embedded/test contexts */
function persistedTrackChanges(): boolean {
  try {
    return localStorage.getItem(TRACK_CHANGES_KEY) === '1'
  } catch {
    return false
  }
}

function executionError(exec: ToolExecution): BridgeCommandResult {
  return fail(exec.output === STALE_DOC_ERROR ? 'stale_document' : 'invalid_params', exec.output)
}

/**
 * Build the per-tab bridge command handler. The handler owns one slot of
 * bridge-undo state; App.tsx keeps a single instance alive per document.
 */
export function createBridgeCommandHandler(
  deps: BridgeCommandDeps,
): (method: string, params: Record<string, unknown>) => Promise<BridgeCommandResult> {
  let lastTurn: BridgeTurn | null = null

  const track = (): { author: string } | undefined =>
    (deps.isTrackChangesOn?.() ?? persistedTrackChanges())
      ? { author: BRIDGE_REVISION_AUTHOR }
      : undefined

  const numIds = (state: BridgeDocState): NumIds => ({
    bullet: findNumId(state.blocks, 'bullet') ?? (state.isBlank ? BLANK_BULLET_NUM_ID : null),
    ordered: findNumId(state.blocks, 'ordered') ?? (state.isBlank ? BLANK_ORDERED_NUM_ID : null),
  })

  /** run one agent tool by name against the live editor, freezing this call's selection */
  const runTool = async (
    editor: Editor,
    state: BridgeDocState,
    name: string,
    input: Record<string, unknown>,
  ): Promise<ToolExecution> => {
    // fresh per call: the bridge client is slower than a human, so a selection
    // frozen at get_context time must not silently survive user clicks
    const frozen = { scope: getSelectionScope(editor), doc: editor.state.doc }
    return executeTool(
      editor,
      { id: `bridge-${name}`, name, input },
      numIds(state),
      track(),
      undefined,
      frozen,
      deps.getComments?.(),
      deps.getHf?.(),
    )
  }

  return async (method, params) => {
    if (
      method !== 'get_context' &&
      method !== 'apply_ops' &&
      method !== 'insert_content' &&
      method !== 'undo'
    ) {
      return fail('unknown_method', `unknown bridge method "${method}"`)
    }
    const editor = deps.getEditor()
    const state = deps.getDocState()
    if (!editor || !state) {
      return fail('no_active_document', 'this tab has no open document')
    }

    if (method === 'get_context') {
      // the context the client receives is the freshness baseline for its
      // next index-addressed write (same contract as the agent's buildContext)
      markDocSeen(editor)
      const context = buildDocContext(
        editor,
        getSelectionScope(editor),
        deps.getComments?.()?.list(),
        deps.getHf?.()?.read(),
      )
      return { ok: true, result: { context, filePath: state.filePath } }
    }

    if (method === 'undo') {
      const turn = lastTurn
      if (!turn) return fail('nothing_to_undo', 'no bridge turn to undo yet')
      if (editor.state.doc !== turn.afterDoc) {
        return fail(
          'stale_document',
          'the document changed since the last bridge turn; undoing would discard those edits — fetch fresh context before editing',
        )
      }
      const run = editor
        .chain()
        .setMeta(TABLE_TRAILING_SKIP, true)
        .setContent(turn.before as never)
        .run()
      lastTurn = null
      if (!run) return fail('internal', 'undo failed to apply the pre-turn snapshot')
      // the rewound doc is the new freshness baseline for the client
      markDocSeen(editor)
      return { ok: true, result: { undone: true } }
    }

    // mutating commands: snapshot first so one undo reverts the whole turn
    const before = editor.getJSON()
    const exec =
      method === 'insert_content'
        ? await runTool(editor, state, 'insert_content', {
            html: String(params.html ?? ''),
            ...(params.afterBlockIndex !== undefined
              ? { afterBlockIndex: params.afterBlockIndex }
              : {}),
          })
        : await runTool(editor, state, 'apply_ops', { ops: params.ops })
    if (exec.isError) return executionError(exec)
    if (exec.mutated) lastTurn = { before, afterDoc: editor.state.doc }
    return { ok: true, result: exec.output }
  }
}
