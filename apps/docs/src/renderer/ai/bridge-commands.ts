// Live-bridge command handlers (Airy Copilot): executes bridge requests from
// the shell's local socket server against THIS tab's editor by reusing the
// embedded agent pipeline — buildDocContext for context+selection, executeTool
// for insert_content/apply_ops (one ProseMirror transaction each, same
// stale-guard and tracked-changes behavior), and the AI panel's snapshot
// rollback for undo. One bridge call = one "turn"; turns sit on a small LIFO
// stack, undo reverts exactly the latest turn (the MCP live_apply_ops contract
// "each bridge call is one undo step" — a combined html+ops call pushes two
// turns, so live_undo reverts it in two steps) and refuses when the user has
// edited since. Turns are stamped with the calling bridge connection's id: the
// automatic rollback after a failed turn (params.ownTurnsOnly) only reverts
// the requester's OWN turn — with two copilot clients on one document it must
// never silently revert the other client's edit (that is a
// turn_owned_by_other error instead; an explicit undo stays allowed and
// reports whose turn it reverted).
import type { Editor, JSONContent } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID, type Block } from '@airy-office/docx-engine'

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
  /** id of the bridge connection that made the turn (turn ownership) */
  owner: string
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
 * Build the per-tab bridge command handler. The handler owns the bridge-undo
 * turn stack; App.tsx keeps a single instance alive per document.
 *
 * `clientId` is the calling bridge connection's identity (stamped by the
 * shell's bridge server, one per socket). Callers without one — the embedded
 * AI panel context, tests, older preloads — share the 'unknown' identity,
 * which matches the previous single-client behavior.
 */
export function createBridgeCommandHandler(
  deps: BridgeCommandDeps,
): (
  method: string,
  params: Record<string, unknown>,
  clientId?: string,
) => Promise<BridgeCommandResult> {
  // LIFO of bridge turns. A single slot could not hold a combined
  // live_apply_ops (the MCP server sends insert_content and apply_ops as two
  // bridge calls): the ops turn overwrote the insert turn, the first undo
  // reverted only the ops and the second found nothing — the insert was
  // stranded (BUG-1505). Each undo pops one turn, so the combined call takes
  // exactly the two undos the live_apply_ops contract documents. Deeper turns
  // stay reachable (a later failed call never orphans an earlier one) but the
  // stack is bounded: each turn holds a full document snapshot, and undoing
  // under newer edits is refused by the stale guard anyway.
  const turns: BridgeTurn[] = []
  const MAX_BRIDGE_TURNS = 16
  const connectionId = (clientId?: string): string =>
    typeof clientId === 'string' && clientId !== '' ? clientId : 'unknown'

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

  return async (method, params, clientId) => {
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
      const turn = turns[turns.length - 1]
      if (!turn) return fail('nothing_to_undo', 'no bridge turn to undo yet')
      // content comparison, not identity: undoing the turn above this one
      // rebuilds the doc node, so a turn reached through the stack is only
      // current by CONTENT (eq is ProseMirror's structural equality)
      if (!editor.state.doc.eq(turn.afterDoc)) {
        return fail(
          'stale_document',
          'the document changed since the last bridge turn; undoing would discard those edits — fetch fresh context before editing',
        )
      }
      const caller = connectionId(clientId)
      // the automatic rollback after a failed turn may only revert the
      // requester's OWN turn: with several copilot clients connected, the
      // last turn may belong to another client and reverting it silently
      // would destroy their edit while reporting "back at its pre-call state"
      if (params.ownTurnsOnly === true && turn.owner !== caller) {
        return fail(
          'turn_owned_by_other',
          `the last bridge turn belongs to another copilot client (${turn.owner}); the automatic rollback refuses to revert it — undo without ownTurnsOnly to revert it intentionally`,
        )
      }
      const run = editor
        .chain()
        .setMeta(TABLE_TRAILING_SKIP, true)
        .setContent(turn.before as never)
        .run()
      // pop only when the rollback applied — a failed setContent leaves the
      // turn in place so the undo can be retried
      if (!run) return fail('internal', 'undo failed to apply the pre-turn snapshot')
      turns.pop()
      // the rewound doc is the new freshness baseline for the client
      markDocSeen(editor)
      // an explicit undo may legitimately revert another client's turn (the
      // user asked for it) — the result says whose turn it was
      const revertedOther = turn.owner !== caller
      return {
        ok: true,
        result: {
          undone: true,
          ...(revertedOther ? { revertedTurnOf: turn.owner, anotherClient: true } : {}),
        },
      }
    }

    // mutating commands: snapshot first so one undo reverts the whole turn
    const before = editor.getJSON()
    const exec =
      method === 'insert_content'
        ? await runTool(editor, state, 'insert_content', {
            html: String(params.html ?? ''),
            // the bridge contract (live_apply_ops) inserts at the END of the
            // document so block indexes from get_context stay valid — the
            // embedded pipeline's cursor default only applies to panel
            // callers, which pass their own afterBlockIndex
            ...(params.afterBlockIndex !== undefined
              ? { afterBlockIndex: params.afterBlockIndex }
              : { afterBlockIndex: editor.state.doc.childCount - 1 }),
          })
        : await runTool(editor, state, 'apply_ops', { ops: params.ops })
    if (exec.isError) return executionError(exec)
    if (exec.mutated) {
      turns.push({ before, afterDoc: editor.state.doc, owner: connectionId(clientId) })
      if (turns.length > MAX_BRIDGE_TURNS) turns.shift()
    }
    return { ok: true, result: exec.output }
  }
}
