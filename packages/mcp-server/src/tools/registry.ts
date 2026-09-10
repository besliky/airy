// Single registration point for every MCP tool (ADR-3): all registerTool calls
// live in this module, isolating the rest of the server from the SDK tool API
// so a future SDK migration (v1 -> v2) only has to touch this file.
//
// Since S6 the document tools are multi-format: open_document accepts
// .docx/.xlsx/.xlsm/.xls/.ods/.doc/.odt and hands out handles from one shared
// session store; read_document stays the text-document reader (docx sessions
// and read-only text), while workbooks are read through read_workbook — a
// separate tool rather than a read_document extension because the addressing
// models differ fundamentally (block indexes vs sheet + A1 ranges), and one
// zod schema cannot describe both without confusing agents.
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'

import { DocxSession, type SessionMeta } from '../docx/session.js'
import { opSignatures, type Op } from '../docx/ops.js'
import { openDocument } from '../import/open.js'
import { BridgeClientError, sharedLiveBridge } from '../live/client.js'
import { getSession, removeSession, storeSession, type DocumentSession } from '../sessions/store.js'
import { TextSession } from '../sessions/text.js'
import { XlsxSession, type XlsxSessionMeta } from '../xlsx/session.js'
import { SERVER_NAME } from '../version.js'

/** text + mirrored JSON content, the repo's standard tool output shape */
function content(payload: object, text?: string) {
  const structuredContent = payload as Record<string, unknown>
  return {
    content: [{ type: 'text' as const, text: text ?? JSON.stringify(payload) }],
    structuredContent,
  }
}

const OPS_GUIDE = [
  'Each op is a flat record { op, target?, ...fields }; fields are patches: present = set, null = clear, absent = keep.',
  'Target conditions (AND, at least one): nodeType ("heading"|"paragraph"|"listItem"|"image"), headingLevel (1-6), containsText (+ matchCase: false), blockIndexes[].',
  ...opSignatures().map((s) => `- ${s}`),
].join('\n')

export function registerTools(server: McpServer): void {
  // Liveness probe: lets an agent confirm the server is reachable before the
  // real document tools arrive in later phases.
  server.registerTool(
    'ping',
    {
      title: 'Ping',
      description: 'Check that the Airy MCP server is reachable and responsive',
      // Empty zod raw shape: the tool takes no arguments
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
      },
    },
    async () => {
      const structuredContent = {
        pong: true,
        server: SERVER_NAME,
        time: new Date().toISOString(),
      }
      // Structured output SHOULD also be mirrored as serialized JSON in a text
      // content block (MCP spec, Server Tools)
      return {
        content: [{ type: 'text', text: JSON.stringify(structuredContent) }],
        structuredContent,
      }
    },
  )

  // ---- headless document editing (Phase 1 docx, Phase 3/3b xlsx + legacy) ----

  server.registerTool(
    'open_document',
    {
      title: 'Open document',
      description:
        'Open a document and return a session handle for the other document tools. Supported ' +
        'formats: .docx and .xlsx/.xlsm open natively (fully editable); .xls and .ods import via ' +
        'conversion (editable as .xlsx; styling is lost — see warnings); .doc and .odt convert to ' +
        '.docx via LibreOffice when installed (editable; save_document format "origin" exports ' +
        'back); without LibreOffice a .doc still opens read-only as extracted text (editable: ' +
        'false). The path must be absolute or workspace-relative and stay inside the server ' +
        'workspace root (AIRY_WORKSPACE_ROOT env var, default: the process working directory). ' +
        'Read-only: nothing is written until save_document. Close sessions with close_document.',
      inputSchema: {
        path: z
          .string()
          .min(1)
          .describe('Absolute or workspace-relative path of the document to open'),
      },
      annotations: {
        readOnlyHint: true,
      },
    },
    async ({ path }) => {
      const session = await openDocument(path)
      storeSession(session)
      const meta = session.meta()
      return content(meta, summarizeOpenMeta(meta))
    },
  )

  server.registerTool(
    'read_document',
    {
      title: 'Read document',
      description:
        'Read an open text document (.docx sessions, or read-only text sessions from legacy ' +
        '.doc). By default returns the block overview ("index|type|content preview" one line per ' +
        'block, plus full-text word/character stats). Pass blocks (indexes) or range ({start,end}) ' +
        'to get the full content of those blocks as restricted HTML (p, h1-h6, ul/ol/li, ' +
        'strong/em/u/s, a, br, table). Block indexes are the addressing scheme for insert_content ' +
        '(at) and apply_ops targets; re-read after edits — indexes shift. For workbook (.xlsx) ' +
        'sessions use read_workbook instead.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        blocks: z
          .array(z.number().int().min(0))
          .max(200)
          .optional()
          .describe('Block indexes to return in full (restricted HTML)'),
        range: z
          .object({ start: z.number().int().min(0), end: z.number().int().min(0) })
          .optional()
          .describe('Inclusive block range to return in full (alternative to blocks)'),
      },
      annotations: {
        readOnlyHint: true,
      },
    },
    async ({ handle, blocks, range }) => {
      const session = getSession(handle)
      if (session instanceof XlsxSession) {
        throw new Error(
          'This handle is a workbook session; use read_workbook (sheet + range) instead.',
        )
      }
      if (session instanceof TextSession) {
        return { content: [{ type: 'text' as const, text: session.readDocument() }] }
      }
      const text = session.readDocument({
        ...(blocks !== undefined ? { blocks } : {}),
        ...(range !== undefined ? { range } : {}),
      })
      return { content: [{ type: 'text' as const, text }] }
    },
  )

  server.registerTool(
    'read_workbook',
    {
      title: 'Read workbook',
      description:
        'Read an open workbook (.xlsx/.xlsm/.xls/.ods sessions). Without options returns the ' +
        'sheet overview ("index|name|id|rows x cols", one line per sheet). With sheet (name or ' +
        'index) and an A1-style range ("A1:E10", single cell "B2") returns a pipe table of cell ' +
        'values; formula cells render as "=FORMULA (cached value)". Without a range the sheet\'s ' +
        'top-left corner (up to 20 rows x 10 columns) is returned as a starting point. Re-read ' +
        'after save_document — sessions always reflect the latest save.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        sheet: z
          .union([z.string().min(1), z.number().int().min(0)])
          .optional()
          .describe('Sheet name or 0-based index (from the overview)'),
        range: z
          .string()
          .min(1)
          .optional()
          .describe('A1-style range within the sheet, e.g. "A1:E10" or "B2"'),
      },
      annotations: {
        readOnlyHint: true,
      },
    },
    async ({ handle, sheet, range }) => {
      const session = getSession(handle)
      if (!(session instanceof XlsxSession)) {
        throw new Error('This handle is not a workbook session; use read_document instead.')
      }
      const text = await session.readWorkbook({
        ...(sheet !== undefined ? { sheet } : {}),
        ...(range !== undefined ? { range } : {}),
      })
      return { content: [{ type: 'text' as const, text }] }
    },
  )

  server.registerTool(
    'insert_content',
    {
      title: 'Insert content',
      description:
        'Insert new content into an open .docx document from a restricted HTML fragment (no DOM ' +
        'features needed). Supported tags: p, h1-h6, ul, ol, li (nested lists allowed), strong/b, ' +
        'em/i, u, s, a[href], br, blockquote, pre, table/tr/th/td (header row styled, cells plain ' +
        'text). Unknown tags keep their text; markdown fences and plain text are tolerated (blank ' +
        'lines split paragraphs). Insertion happens in memory; persist with save_document.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        html: z.string().min(1).describe('Restricted HTML fragment to insert'),
        at: z
          .number()
          .int()
          .min(-1)
          .optional()
          .describe(
            'Insert after this block index (-1 = document start; default: end of document)',
          ),
      },
      annotations: {
        destructiveHint: false,
      },
    },
    async ({ handle, html, at }) => {
      const session = requireDocxSession(handle)
      const { inserted } = session.insertContent(html, at ?? Number.MAX_SAFE_INTEGER)
      const meta = session.meta()
      return content(
        {
          inserted,
          blockCount: meta.blockCount,
          dirty: meta.dirty,
        },
        `Inserted ${inserted} block(s). Subsequent block indexes have shifted; call read_document if you need the new state.`,
      )
    },
  )

  server.registerTool(
    'apply_ops',
    {
      title: 'Apply edit operations',
      description:
        'Apply a batch of canonical edit operations to an open .docx document. The batch is ' +
        'validated up front and applied atomically: any invalid op rejects the whole batch with ' +
        'an error and nothing is applied. ' +
        `Operations:\n${OPS_GUIDE}\nEdits happen in memory; persist with save_document.`,
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        ops: z
          .array(z.record(z.string(), z.unknown()))
          .min(1)
          .max(100)
          .describe('Batch of op records, applied in order'),
        dryRun: z
          .boolean()
          .optional()
          .describe('Validate and report the batch without applying anything'),
      },
      annotations: {
        destructiveHint: false,
      },
    },
    async ({ handle, ops, dryRun }) => {
      const session = requireDocxSession(handle)
      const { results, summary, dryRun: isDry } = session.applyOps(ops as Op[], dryRun === true)
      return content(
        { results, summary, dryRun: isDry },
        `${summary}${isDry ? ' (dry run, nothing applied)' : ''}`,
      )
    },
  )

  server.registerTool(
    'save_document',
    {
      title: 'Save document',
      description:
        'Save an open document to disk atomically (temp file + rename). Without a path: a .docx ' +
        'session overwrites the file it was opened from (refusing with a clear error when the ' +
        'file changed on disk since open — reopen and re-apply in that case), while sessions ' +
        'imported from .xls/.ods/.doc/.odt write a fresh sibling file with the native extension ' +
        'next to the original (the original stays untouched). format "origin" instead exports ' +
        'back to the original .doc/.odt/.ods through LibreOffice (best-effort; .xls output is ' +
        'not supported — use the default .xlsx save). Untouched parts of the document are kept ' +
        'byte-identical; a save with zero edits writes the original bytes back verbatim. Returns ' +
        'the absolute path.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        path: z
          .string()
          .min(1)
          .optional()
          .describe('Optional save-as path (workspace-confined); default: see description'),
        format: z
          .enum(['docx', 'xlsx', 'origin'])
          .optional()
          .describe(
            'Output format: default matches the session (docx sessions -> "docx", workbooks -> ' +
              '"xlsx"); "origin" exports back to the original legacy/ODF format via LibreOffice',
          ),
      },
      annotations: {
        destructiveHint: true,
      },
    },
    async ({ handle, path, format }) => {
      const session = getSession(handle)
      if (session instanceof TextSession) {
        throw new Error(
          'This is a read-only text session (legacy .doc without LibreOffice); it cannot be saved. ' +
            'Install LibreOffice to open the document as an editable converted .docx session.',
        )
      }
      if (session instanceof XlsxSession) {
        if (format === 'docx') {
          throw new Error(
            'format "docx" is not valid for a workbook session; use "xlsx" or "origin".',
          )
        }
        const result = await session.save(path, format === 'origin' ? 'origin' : 'xlsx')
        return content(
          result,
          `Saved ${result.bytes} bytes to ${result.path} (${result.format})` +
            `${result.unchanged ? ' — no changes: bytes round-tripped verbatim' : ''}` +
            `${result.warnings.length > 0 ? `. ${result.warnings.join(' ')}` : ''}`,
        )
      }
      if (format === 'xlsx') {
        throw new Error(
          'format "xlsx" is not valid for a text document session; use "docx" or "origin".',
        )
      }
      const result = await session.save(path, format === 'origin' ? 'origin' : 'docx')
      return content(
        result,
        `Saved ${result.bytes} bytes to ${result.path}${result.unchanged ? ' (no changes: bytes round-tripped verbatim)' : ''}`,
      )
    },
  )

  server.registerTool(
    'close_document',
    {
      title: 'Close document',
      description:
        'Close an open session and release its resources: workbook sessions close their sidecar ' +
        'session and remove conversion temp files; sessions converted from .doc/.odt remove their ' +
        'temp .docx. After closing, the handle is invalid (open_document again to continue). ' +
        'Unsaved edits are discarded.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
      },
      annotations: {
        destructiveHint: true,
      },
    },
    async ({ handle }) => {
      const session = getSession(handle)
      const kind = session.meta().kind
      const cleanedTempDirs = await session.close()
      removeSession(handle)
      return content(
        { closed: true, handle, kind, cleanedTempDirs },
        `Closed ${kind} session ${handle}${cleanedTempDirs.length > 0 ? ` (removed ${String(cleanedTempDirs.length)} temp dir(s))` : ''}.`,
      )
    },
  )

  // ---- live bridge: edit the document open in the running app (Phase 2) ----

  server.registerTool(
    'live_status',
    {
      title: 'Live status',
      description:
        'Check the live bridge to the Airy/GenOffice desktop app. Returns {running:false} (no error) when ' +
        'the app is not running — use the headless document tools in that case. When running, returns the ' +
        'bridge protocol version, the app pid, and the list of open documents ({id, title, filePath, active}) ' +
        'of which the live tools always target the active one.',
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
      },
    },
    async () => {
      const bridge = sharedLiveBridge()
      try {
        const pong = asRecord(await bridge.call('ping')) ?? {}
        let documents: unknown[] = []
        try {
          const list = asRecord(await bridge.call('list'))
          if (Array.isArray(list?.documents)) documents = list.documents
        } catch {
          // ping proved the bridge is up; a failing list must not flip running
        }
        const payload = {
          running: true,
          pid: pong.pid,
          protocolVersion: pong.protocolVersion,
          documents,
        }
        return content(
          payload,
          `Live bridge running (pid ${String(pong.pid)}), ${documents.length} open document(s).`,
        )
      } catch (err) {
        const reason = describeBridgeFailure(err)
        return content({ running: false, reason }, reason)
      }
    },
  )

  server.registerTool(
    'live_get_context',
    {
      title: 'Get live document context',
      description:
        'Read the context of the ACTIVE document in the running Airy/GenOffice app: block list (index|type| ' +
        'content preview), the current selection (<sel>), comments, and the file path. The context is the ' +
        'freshness baseline for index-addressed live_apply_ops — refetch it after the user edits or a ' +
        'stale_document error.',
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
      },
    },
    async () => {
      const bridge = sharedLiveBridge()
      let result: unknown
      try {
        result = await bridge.call('get_context')
      } catch (err) {
        throw new Error(describeBridgeFailure(err), { cause: err })
      }
      const record = asRecord(result)
      if (!record) throw new Error('live bridge error: get_context returned a malformed payload')
      return content(record)
    },
  )

  server.registerTool(
    'live_apply_ops',
    {
      title: 'Apply edits to the live document',
      description:
        'Edit the ACTIVE document in the running Airy/GenOffice app in one go: insert a restricted-HTML ' +
        'fragment and/or apply canonical edit ops. When both are given the html is inserted first (at the ' +
        'end of the document, so block indexes from live_get_context stay valid) and the ops then run against ' +
        'the result — a single call can add a section and format it. The user sees the change immediately; ' +
        'tracked changes are authored as "Airy Copilot" when the app has track changes on. Each bridge call ' +
        'is one undo step, so undo a combined edit with live_undo twice. ' +
        `Operations:\n${OPS_GUIDE}`,
      inputSchema: {
        ops: z
          .array(z.record(z.string(), z.unknown()))
          .min(1)
          .max(100)
          .optional()
          .describe('Batch of op records, applied in order'),
        html: z
          .string()
          .min(1)
          .optional()
          .describe('Restricted HTML fragment to insert at the end of the document'),
      },
      annotations: {
        destructiveHint: true,
      },
    },
    async ({ ops, html }) => {
      if (ops === undefined && html === undefined) {
        throw new Error('live_apply_ops requires at least one of ops or html')
      }
      const bridge = sharedLiveBridge()
      const applied: Record<string, unknown> = {}
      try {
        if (html !== undefined) applied.insert = await bridge.call('insert_content', { html })
        if (ops !== undefined) applied.ops = await bridge.call('apply_ops', { ops })
      } catch (err) {
        throw new Error(describeBridgeFailure(err), { cause: err })
      }
      const parts = [
        ...(applied.insert !== undefined ? ['inserted content'] : []),
        ...(applied.ops !== undefined ? [`applied ${String(ops?.length ?? 0)} op(s)`] : []),
      ]
      return content(applied, `Applied to the live document: ${parts.join(' and ')}.`)
    },
  )

  server.registerTool(
    'live_undo',
    {
      title: 'Undo last live edit',
      description:
        'Revert the last live bridge turn in the active document of the running Airy/GenOffice app (one ' +
        'live_apply_ops / live_undo step). Refuses with nothing_to_undo when the agent made no edits yet, ' +
        'and with stale_document when the user edited the document since — fetch fresh context instead.',
      inputSchema: {},
      annotations: {
        destructiveHint: true,
      },
    },
    async () => {
      const bridge = sharedLiveBridge()
      let result: unknown
      try {
        result = await bridge.call('undo')
      } catch (err) {
        throw new Error(describeBridgeFailure(err), { cause: err })
      }
      return content(asRecord(result) ?? { undone: true }, 'Undid the last live bridge turn.')
    },
  )
}

// ---- session helpers ----

/** Docx-only tools (insert_content/apply_ops) reject other session kinds clearly. */
function requireDocxSession(handle: string): DocxSession {
  const session: DocumentSession = getSession(handle)
  if (!(session instanceof DocxSession)) {
    throw new Error(
      'insert_content/apply_ops are only available for editable .docx sessions ' +
        `(this handle is a "${session.meta().kind}" session).`,
    )
  }
  return session
}

// ---- live bridge helpers ----

function asRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null
}

/** agent-facing message for any live bridge failure */
function describeBridgeFailure(err: unknown): string {
  if (err instanceof BridgeClientError) {
    if (err.code === 'bridge_not_running') {
      return (
        `live bridge not running: ${err.message}. ` +
        'Start the GenOffice (Airy) desktop app with a document open, or point AIRY_BRIDGE_FILE at its ' +
        'airy-bridge.json, then retry.'
      )
    }
    if (err.code === 'bridge_unauthorized') {
      return (
        `live bridge unauthorized: ${err.message}. ` +
        'The app may have restarted and issued a fresh token — the next call rereads it automatically; ' +
        'verify AIRY_BRIDGE_FILE if it persists.'
      )
    }
    return `live bridge error${err.bridgeCode ? ` (${err.bridgeCode})` : ''}: ${err.message}`
  }
  return err instanceof Error ? err.message : String(err)
}

// ---- open_document summaries (one per session kind) ----

type AnyOpenMeta = SessionMeta | XlsxSessionMeta | ReturnType<TextSession['meta']>

function summarizeOpenMeta(meta: AnyOpenMeta): string {
  if (meta.kind === 'xlsx') {
    const names = meta.sheets.map((sheet) => sheet.name).join(', ')
    const suffix = meta.converted
      ? ` (imported from .${meta.format} — ${meta.warnings[0] ?? 'conversion'})`
      : ''
    return (
      `Opened ${meta.path.split('/').pop() ?? meta.path} as an editable workbook${suffix}: ` +
      `${String(meta.sheets.length)} sheet(s) — ${names}. Handle: ${meta.handle}. Path: ${meta.path}`
    )
  }
  if (meta.kind === 'text') {
    return (
      `Opened ${meta.fileName} read-only (.${meta.format}, text extraction): ${String(meta.wordCount)} words, ` +
      `${String(meta.charCount)} characters. ${meta.warnings[0] ?? ''} Handle: ${meta.handle}. Path: ${meta.path}`
    )
  }
  return (
    `Opened ${meta.fileName}: ${meta.blockCount} blocks, ${meta.wordCount} words ` +
    `(${meta.charCount} characters). Handle: ${meta.handle}. Path: ${meta.path}`
  )
}
