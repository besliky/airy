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
// zod schema cannot describe both without confusing agents. Since PAR-003,
// .md/.markdown open as line-based text sessions: read_document shows the
// heading structure plus the text (blocks/range address lines), insert_content
// takes markdown text at marker/heading/line positions, and apply_ops runs a
// line-op vocabulary on markdown sessions. Since PAR-004, .html/.htm do the
// same with a parse5 structure summary (headings, links, title) and verbatim
// HTML fragment inserts.
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { z } from 'zod'

import { DocxSession, type SessionMeta } from '../docx/session.js'
import { opSignatures, type Op } from '../docx/ops.js'
import { openDocument } from '../import/open.js'
import { HtmlSession } from '../html/session.js'
import { BridgeClientError, sharedLiveBridge } from '../live/client.js'
import { MarkdownSession } from '../markdown/session.js'
import { getSession, removeSession, storeSession, type DocumentSession } from '../sessions/store.js'
import { TextSession } from '../sessions/text.js'
import { XlsxSession, type XlsxSessionMeta } from '../xlsx/session.js'
import { richRunSchema, workbookStyleEditSchema } from '../xlsx/save.js'
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
  'Target conditions (AND, at least one): nodeType ("heading"|"paragraph"|"listItem"|"image"; the renderer spellings "docHeading"|"docParagraph"|"docListItem" are accepted aliases), headingLevel (1-6), containsText (+ matchCase: false), blockIndexes[].',
  ...opSignatures().map((s) => `- ${s}`),
].join('\n')

/** line-op vocabulary html sessions accept in apply_ops */
const HTML_OPS_GUIDE = [
  'HTML sessions accept a line-op vocabulary instead (no target; line indexes are 0-based and shift after every splice, so re-read between batches):',
  '- insertLines after(-1 = start) text — splice HTML/markup source after a line',
  '- replaceLines from to text — replace an inclusive line range (empty text deletes the range)',
  '- deleteLines from to — remove an inclusive line range',
  '- findReplace find replace matchCase?(default true) from? to? — line-scoped replace, optionally within an inclusive line window',
].join('\n')

/** line-op vocabulary markdown sessions accept in apply_ops */
const MARKDOWN_OPS_GUIDE = [
  'Markdown sessions accept a line-op vocabulary instead (no target; line indexes are 0-based and shift after every splice, so re-read between batches):',
  '- insertLines after(-1 = start) text — splice markdown source after a line',
  '- replaceLines from to text — replace an inclusive line range (empty text deletes the range)',
  '- deleteLines from to — remove an inclusive line range',
  '- findReplace find replace matchCase?(default true) from? to? — line-scoped replace, optionally within an inclusive line window',
].join('\n')

/** one agent-issued cell edit: sheet + A1 ref + value/formula/style patches */
const CELL_EDIT_SCHEMA = z
  .object({
    sheet: z
      .union([z.string().min(1), z.number().int().min(0)])
      .describe('Sheet name or 0-based index (from the read_workbook overview)'),
    ref: z.string().min(1).max(16).describe('Single cell in A1 notation, e.g. "B2"'),
    value: z
      .union([z.string().max(32_767), z.number().finite(), z.boolean(), z.null()])
      .optional()
      .describe('Constant to store (null clears the value); ignored when formula is given'),
    formula: z
      .string()
      .min(1)
      .max(8_192)
      .optional()
      .describe(
        'Formula, with or without the leading "="; written without a cached result so ' +
          'apps recalculate on open',
      ),
    style: workbookStyleEditSchema
      .optional()
      .describe(
        'Style patch (bold, italic, fillColor "#RRGGBB", fontColor, numberFormat, ' +
          'horizontalAlignment, borders, ...): only the given keys change, the rest of the ' +
          "cell's format survives",
      ),
    rich: z
      .array(richRunSchema)
      .max(1_000)
      .optional()
      .describe(
        'Rich-text runs for a string value ({ text, bold, italic, underline, strikethrough, ' +
          'color?, size?, family?, vertAlign? }); the joined run text becomes the cell value',
      ),
    styleReset: z
      .boolean()
      .optional()
      .describe('Reset the cell to the default style before applying the style patch'),
  })
  .strict()
  .refine(
    (edit) =>
      edit.value !== undefined ||
      edit.formula !== undefined ||
      edit.style !== undefined ||
      edit.rich !== undefined ||
      edit.styleReset !== undefined,
    { message: 'A cell edit needs at least one of value, formula, style, rich or styleReset.' },
  )

/** ops the embedded live registry accepts beyond the headless guide (kept in
 * sync with the apps/docs/src/renderer/ai/ops.ts signatures) */
const LIVE_OPS_EXTRAS = [
  'Live-only extras accepted by the embedded registry in addition to the ops above:',
  '- setImageProperties <target nodeType "image"> widthPx? heightPx? align? ("left"|"center"|"right"|null) — image blocks only; giving one dimension scales the other proportionally',
  '- insertToc afterBlockIndex (-1 = document start) — insert a TOC field built from the current headings; Word computes page numbers on open (fails when the document has no headings)',
  '- setFont / setMatchedFont additionally accept link: { url } | null over this bridge.',
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
        'formats: .docx and .xlsx/.xlsm open natively (docx fully editable via insert_content/' +
        'apply_ops; workbooks editable in cell values, formulas and styles via apply_workbook_ops ' +
        '— charts, pivots and sheet structure are not editable headlessly); .xls and .ods import via ' +
        'conversion (editable as .xlsx; styling is lost — see warnings); .doc and .odt convert to ' +
        '.docx via LibreOffice when installed (editable; save_document format "origin" exports ' +
        'back); without LibreOffice a .doc still opens read-only as extracted text (editable: ' +
        'false). .md/.markdown open as line-based text sessions (read_document shows heading ' +
        'structure plus text; insert_content inserts markdown source; apply_ops runs line ops; ' +
        'UTF-8 with BOM/EOL preservation — files that are not valid UTF-8 are refused), and ' +
        '.html/.htm likewise (read_document shows a parse5 structure summary — ' +
        'headings/links/title; insert_content splices an HTML fragment verbatim; declared legacy ' +
        'charsets accepted, undeclared non-UTF-8 refused). The path ' +
        'must be absolute or workspace-relative and stay inside the server ' +
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
        'Read an open text document (.docx, markdown, or html sessions, or read-only text ' +
        'sessions from legacy .doc). By default returns the block overview ("index|type|content ' +
        'preview" one line per block, plus full-text word/character stats). Pass blocks (indexes) ' +
        'or range ({start,end}) to get the full content of those blocks as restricted HTML (p, ' +
        'h1-h6, ul/ol/li, strong/em/u/s, a, br, table). Block indexes are the addressing scheme ' +
        'for insert_content (at) and apply_ops targets; re-read after edits — indexes shift. ' +
        'Markdown sessions: the default read shows stats (EOL/BOM included), the heading list ' +
        '"ordinal|line|level|text" (capped at 200 entries) and the full text; the whole read is ' +
        'truncated at 30k characters — narrow with blocks/range, which address LINES. HTML ' +
        'sessions: the default read adds the title and a parse5 structure summary — headings ' +
        '"ordinal|line|level|text" and links "ordinal|line|text -> href" with line positions ' +
        '(both capped at 200 entries, counted toward the 30k budget). ' +
        'For workbook (.xlsx) sessions use ' +
        'read_workbook instead.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        blocks: z
          .array(z.number().int().min(0))
          .max(200)
          .optional()
          .describe('Block indexes to return in full (lines, for markdown/html sessions)'),
        range: z
          .object({ start: z.number().int().min(0), end: z.number().int().min(0) })
          .optional()
          .describe(
            'Inclusive block/line range to return in full (alternative to blocks; ' +
              'spans cap at 10,000 — split larger ranges)',
          ),
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
      if (session instanceof MarkdownSession || session instanceof HtmlSession) {
        const text = session.readDocument({
          ...(blocks !== undefined ? { lines: blocks } : {}),
          ...(range !== undefined ? { range } : {}),
        })
        return { content: [{ type: 'text' as const, text }] }
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
    'apply_workbook_ops',
    {
      title: 'Apply workbook edits',
      description:
        'Apply a batch of cell edits to an open workbook session (.xlsx/.xlsm/.xls/.ods). Each ' +
        'edit targets one cell by sheet (name or index) and A1 ref and may set a value (string / ' +
        'number / boolean / null), a formula (stored without a cached result, so spreadsheet apps ' +
        'recalculate on open), a style patch (bold, fillColor "#RRGGBB", numberFormat, borders, ' +
        '...), rich-text runs, or a combination. Fields are patches: a value/formula edit replaces ' +
        "the cell's content, a style edit replaces only the cell's format, and later edits to the " +
        'same cell win per channel. The batch is validated up front (unknown sheets, bad refs, ' +
        'malformed edits reject the whole batch); dryRun reports without journaling. Edits are ' +
        'journaled in memory; persist with save_document, which keeps untouched zip entries ' +
        'byte-identical (xl/workbook.xml excepted — it gains the fullCalcOnLoad recalc flag ' +
        'when the source lacks it). Cell values, formulas and styles only — charts, pivots, merged ranges and ' +
        'sheet structure are not editable headlessly.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        edits: z
          .array(CELL_EDIT_SCHEMA)
          .min(1)
          .max(100)
          .describe('Batch of cell edits, journaled in order'),
        dryRun: z
          .boolean()
          .optional()
          .describe('Validate and report the batch without journaling anything'),
      },
      annotations: {
        destructiveHint: false,
      },
    },
    async ({ handle, edits, dryRun }) => {
      const session = requireWorkbookSession(handle)
      // journal per sheet (first-appearance order keeps the batch's order)
      const bySheet = new Map<string | number, typeof edits>()
      for (const edit of edits) {
        const group = bySheet.get(edit.sheet) ?? []
        group.push(edit)
        bySheet.set(edit.sheet, group)
      }
      const batches = [...bySheet.entries()].map(([sheet, group]) => ({
        sheet,
        cells: group.map(({ sheet: _sheet, ...cell }) => cell),
      }))
      // validate the whole batch up front so a bad edit rejects everything
      // (setCells in dry-run mode parses every sheet name and ref)
      for (const batch of batches) session.setCells(batch, true)
      // report the journal's merged entry count: several edits to one cell
      // collapse into a single journaled entry
      let journaled = 0
      if (dryRun !== true) {
        for (const batch of batches) journaled += session.setCells(batch, false).merged
      } else {
        journaled = edits.length
      }
      const meta = session.meta()
      return content(
        { journaled, dirty: meta.dirty, dryRun: dryRun === true },
        `Journaled ${journaled} cell edit(s)${dryRun === true ? ' (dry run, nothing applied)' : ''}; ` +
          'persist with save_document.',
      )
    },
  )

  server.registerTool(
    'insert_content',
    {
      title: 'Insert content',
      description:
        'Insert new content into an open .docx, markdown, or html document. Docx sessions take a ' +
        'restricted HTML fragment (no DOM features needed). Supported tags: p, h1-h6, ul, ol, li ' +
        '(nested lists allowed), strong/b, em/i, u, s, a[href], br, blockquote, pre, table/tr/th/td ' +
        '(header row styled, cells plain text). Unknown tags keep their text; markdown fences and ' +
        'plain text are tolerated (blank lines split paragraphs). Link policy: a[href] accepts ' +
        'http/https, mailto, #fragment and scheme-less relative hrefs only — anchors with any ' +
        'other scheme (javascript:, file:, data:, …) degrade to plain text. Markdown sessions ' +
        "instead take `text` (markdown source; the file's EOL style and BOM are preserved) at one " +
        'of three positions: after the first line containing `marker`, after heading N ' +
        '(`afterHeading`, 1-based ordinal from the read structure), or after line `at` ' +
        '(-1 = document start; default: end of document; marker > afterHeading > at). HTML ' +
        'sessions take the fragment VERBATIM (no reparse or rewrite — exactly these bytes land ' +
        "on disk, modulo the file's EOL style) at one of two positions: after the first line " +
        'containing `marker` (e.g. "</body>" to append rendered content), or after line `at` ' +
        '(-1 = document start; default: end of document; marker > at). Passing afterHeading to ' +
        'an html session is an explicit error — it is a markdown-session option. Insertion ' +
        'happens in memory; persist with save_document.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        html: z
          .string()
          .min(1)
          .optional()
          .describe(
            'Restricted HTML fragment (docx sessions) or verbatim fragment (html sessions)',
          ),
        text: z
          .string()
          .min(1)
          .optional()
          .describe('Markdown source to insert (markdown sessions)'),
        at: z
          .number()
          .int()
          .min(-1)
          .optional()
          .describe(
            'Insert after this block/line index (-1 = document start; default: end of document)',
          ),
        afterHeading: z
          .number()
          .int()
          .min(1)
          .optional()
          .describe(
            'Markdown sessions: insert after heading N (1-based ordinal from the read structure)',
          ),
        marker: z
          .string()
          .min(1)
          .optional()
          .describe(
            'Markdown/HTML sessions: insert after the first line containing this exact substring ' +
              '(html: e.g. "</body>"); ignored for docx sessions',
          ),
      },
      annotations: {
        destructiveHint: false,
      },
    },
    async ({ handle, html, text, at, afterHeading, marker }) => {
      const session = getSession(handle)
      if (session instanceof MarkdownSession) {
        if (html !== undefined) {
          throw new Error(
            'This handle is a markdown session: pass markdown source in `text`, not `html`.',
          )
        }
        const result = session.insertContent(text ?? '', {
          ...(at !== undefined ? { at } : {}),
          ...(afterHeading !== undefined ? { afterHeading } : {}),
          ...(marker !== undefined ? { marker } : {}),
        })
        return content(
          {
            inserted: result.inserted,
            at: result.at,
            lineCount: result.lineCount,
            dirty: result.dirty,
          },
          `${result.detail}. Subsequent line indexes have shifted; call read_document if you need the new state.`,
        )
      }
      if (session instanceof HtmlSession) {
        if (afterHeading !== undefined) {
          throw new Error(
            'afterHeading is a markdown-session option; html sessions position inserts ' +
              'via marker or at.',
          )
        }
        const result = session.insertContent(html ?? '', {
          ...(at !== undefined ? { at } : {}),
          ...(marker !== undefined ? { marker } : {}),
        })
        return content(
          {
            inserted: result.inserted,
            at: result.at,
            lineCount: result.lineCount,
            dirty: result.dirty,
          },
          `${result.detail}. Subsequent line indexes have shifted; call read_document if you need the new state.`,
        )
      }
      const docx = requireDocxSession(session)
      if (text !== undefined || afterHeading !== undefined || marker !== undefined) {
        throw new Error(
          'text/afterHeading/marker are markdown/html-session insert options; a .docx session ' +
            'takes html (and optionally at).',
        )
      }
      const { inserted } = docx.insertContent(html ?? '', at ?? Number.MAX_SAFE_INTEGER)
      const meta = docx.meta()
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
        'Apply a batch of canonical edit operations to an open .docx, markdown, or html document. ' +
        'The batch is validated up front and applied atomically: any invalid op rejects the whole ' +
        'batch with an error and nothing is applied. ' +
        `Operations:\n${OPS_GUIDE}\n${MARKDOWN_OPS_GUIDE}\n${HTML_OPS_GUIDE}\nEdits happen in memory; persist with save_document.`,
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
      const session = getSession(handle)
      if (session instanceof MarkdownSession || session instanceof HtmlSession) {
        const result = session.applyOps(ops, dryRun === true)
        return content(
          { results: result.results, summary: result.summary, dryRun: result.dryRun },
          `${result.summary}${result.dryRun ? ' (dry run, nothing applied)' : ''}`,
        )
      }
      const docx = requireDocxSession(session)
      const { results, summary, dryRun: isDry } = docx.applyOps(ops as Op[], dryRun === true)
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
        'next to the original (the original stays untouched; when that sibling already exists ' +
        'the save is refused like any save-as target). An explicit path that already ' +
        'exists on disk is refused with an error unless it is a file the session itself opened ' +
        'or saved — pass overwrite: true to replace it. format "origin" instead exports ' +
        'back to the original .doc/.odt/.ods through LibreOffice (best-effort; .xls output is ' +
        'not supported — use the default .xlsx save). Byte preservation differs by format: docx ' +
        'saves keep untouched parts byte-identical and a zero-edit save writes the original bytes ' +
        'back verbatim; markdown and html sessions behave the same at line granularity ' +
        '(untouched lines keep their exact bytes, EOLs included, and a zero-edit save round-trips ' +
        'the file verbatim; an edited save writes UTF-8 with the original BOM flag re-applied, ' +
        'html additionally rewriting a legacy charset declaration to utf-8, and the format ' +
        'parameter is not accepted); xlsx saves keep untouched zip entries ' +
        'byte-identical except ' +
        'xl/workbook.xml, which is rewritten when needed to force recalculation on open (the ' +
        'fullCalcOnLoad flag) — so even a zero-edit xlsx save may touch that one entry, and for ' +
        'workbooks the unchanged result flag reflects the edit journal, not the bytes. Returns ' +
        'the absolute path.',
      inputSchema: {
        handle: z.string().min(1).describe('Session handle from open_document'),
        path: z
          .string()
          .min(1)
          .optional()
          .describe('Optional save-as path (workspace-confined); default: see description'),
        overwrite: z
          .boolean()
          .optional()
          .describe(
            'Allow replacing an existing file at path (default false: an existing unrelated ' +
              'file is refused)',
          ),
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
    async ({ handle, path, overwrite, format }) => {
      const session = getSession(handle)
      if (session instanceof TextSession) {
        throw new Error(
          'This is a read-only text session (legacy .doc without LibreOffice); it cannot be saved. ' +
            'Install LibreOffice to open the document as an editable converted .docx session.',
        )
      }
      if (session instanceof MarkdownSession) {
        if (format !== undefined) {
          throw new Error(
            `format "${format}" is not valid for a markdown session; it always saves UTF-8 markdown.`,
          )
        }
        const saveOptions = overwrite === undefined ? {} : { overwrite }
        const result = await session.save(path, saveOptions)
        return content(
          result,
          `Saved ${result.bytes} bytes to ${result.path}` +
            `${result.unchanged ? ' (no changes: bytes round-tripped verbatim)' : ''}` +
            `${result.warnings.length > 0 ? `. ${result.warnings.join(' ')}` : ''}`,
        )
      }
      if (session instanceof HtmlSession) {
        if (format !== undefined) {
          throw new Error(
            `format "${format}" is not valid for an html session; it always saves UTF-8 html.`,
          )
        }
        const saveOptions = overwrite === undefined ? {} : { overwrite }
        const result = await session.save(path, saveOptions)
        return content(
          result,
          `Saved ${result.bytes} bytes to ${result.path}` +
            `${result.unchanged ? ' (no changes: bytes round-tripped verbatim)' : ''}` +
            `${result.warnings.length > 0 ? `. ${result.warnings.join(' ')}` : ''}`,
        )
      }
      const saveOptions = overwrite === undefined ? {} : { overwrite }
      if (session instanceof XlsxSession) {
        if (format === 'docx') {
          throw new Error(
            'format "docx" is not valid for a workbook session; use "xlsx" or "origin".',
          )
        }
        const result = await session.save(
          path,
          format === 'origin' ? 'origin' : 'xlsx',
          saveOptions,
        )
        return content(
          result,
          `Saved ${result.bytes} bytes to ${result.path} (${result.format})` +
            `${result.unchanged ? ' — no journaled changes: untouched entries preserved (xl/workbook.xml may gain the fullCalcOnLoad recalc flag)' : ''}` +
            `${result.warnings.length > 0 ? `. ${result.warnings.join(' ')}` : ''}`,
        )
      }
      if (format === 'xlsx') {
        throw new Error(
          'format "xlsx" is not valid for a text document session; use "docx" or "origin".',
        )
      }
      const result = await session.save(path, format === 'origin' ? 'origin' : 'docx', saveOptions)
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
        'Check the live bridge to the Airy desktop app. Returns {running:false} (no error) when ' +
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
        'Read the context of the ACTIVE document in the running Airy app: block list (index|type| ' +
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
        'Edit the ACTIVE document in the running Airy app in one go: insert a restricted-HTML ' +
        'fragment and/or apply canonical edit ops. When both are given the html is inserted first (at the ' +
        'end of the document, so block indexes from live_get_context stay valid) and the ops then run against ' +
        'the result — a single call can add a section and format it. When the ops batch fails after the ' +
        'html was inserted, the insert turn is automatically rolled back with an undo, so a failed call ' +
        "leaves the document at its pre-call state; the automatic rollback only reverts THIS client's turn — " +
        'when another copilot client edited the document in between, it refuses (turn_owned_by_other) and the ' +
        'error says the insert remains (call live_undo manually to revert everything). If the rollback undo ' +
        'itself fails, the error says the edit may be partially applied — call live_undo to revert the ' +
        'insert. The user sees the change immediately; ' +
        'tracked changes are authored as "Airy Copilot" when the app has track changes on. Each bridge call ' +
        'is one undo step, so undo a combined edit with live_undo twice. ' +
        `Operations:\n${OPS_GUIDE}\n${LIVE_OPS_EXTRAS}`,
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
      let inserted = false
      try {
        if (html !== undefined) {
          applied.insert = await bridge.call('insert_content', { html })
          inserted = true
        }
        if (ops !== undefined) applied.ops = await bridge.call('apply_ops', { ops })
      } catch (err) {
        if (!inserted) throw new Error(describeBridgeFailure(err), { cause: err })
        // the html landed but the ops batch failed: revert the insert turn so
        // the document is not left half-edited. ownTurnsOnly guards the undo
        // to THIS client's turn — with another copilot client connected, the
        // last turn may be theirs and reverting it silently would destroy
        // their edit while reporting "back at its pre-call state".
        let rollbackNote: string
        try {
          await bridge.call('undo', { ownTurnsOnly: true })
          rollbackNote =
            ' — the inserted html was rolled back with an undo; the document is back at its ' +
            'pre-call state (nothing was applied).'
        } catch (undoErr) {
          if (
            undoErr instanceof BridgeClientError &&
            undoErr.bridgeCode === 'turn_owned_by_other'
          ) {
            rollbackNote =
              ` — the inserted html was NOT rolled back: ${describeBridgeFailure(undoErr)}. ` +
              "The insert is still applied; reverting the other client's turn is a deliberate " +
              'choice — call live_undo manually if that is intended.'
          } else {
            rollbackNote =
              ` — the inserted html could NOT be rolled back (${describeBridgeFailure(undoErr)}); ` +
              'the edit may be partially applied, call live_undo to revert the insert.'
          }
        }
        throw new Error(describeBridgeFailure(err) + rollbackNote, { cause: err })
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
        'Revert the last live bridge turn in the active document of the running Airy app (one ' +
        'live_apply_ops / live_undo step). Refuses with nothing_to_undo when the agent made no edits yet, ' +
        'and with stale_document when the user edited the document since — fetch fresh context instead. ' +
        'When several copilot clients are connected, the last turn may belong to another client: undoing ' +
        'it is allowed (an explicit choice) and the result says whose turn was reverted (anotherClient).',
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
      const record = asRecord(result) ?? { undone: true }
      const note =
        record.anotherClient === true ? ' Undone turn was made by another copilot client.' : ''
      return content(record, `Undid the last live bridge turn.${note}`)
    },
  )
}

// ---- session helpers ----

/** Docx-only editing paths (insert_content/apply_ops docx branch) reject other session kinds clearly. */
function requireDocxSession(session: DocumentSession): DocxSession {
  if (!(session instanceof DocxSession)) {
    throw new Error(
      'insert_content/apply_ops are only available for editable .docx sessions ' +
        `(this handle is a "${session.meta().kind}" session).`,
    )
  }
  return session
}

/** Workbook-only tools (apply_workbook_ops) reject other session kinds clearly. */
function requireWorkbookSession(handle: string): XlsxSession {
  const session: DocumentSession = getSession(handle)
  if (!(session instanceof XlsxSession)) {
    throw new Error(
      'apply_workbook_ops is only available for workbook sessions ' +
        `(this handle is a "${session.meta().kind}" session; ` +
        'use insert_content/apply_ops for text documents).',
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
        'Start the Airy desktop app with a document open, or point AIRY_BRIDGE_FILE at its ' +
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

type AnyOpenMeta =
  | SessionMeta
  | XlsxSessionMeta
  | ReturnType<TextSession['meta']>
  | ReturnType<MarkdownSession['meta']>
  | ReturnType<HtmlSession['meta']>

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
  if (meta.kind === 'markdown') {
    return (
      `Opened ${meta.fileName} as an editable markdown session: ${String(meta.lineCount)} lines, ` +
      `${String(meta.headingCount)} heading(s), ${String(meta.wordCount)} words. ` +
      `${meta.warnings[0] ?? ''}Handle: ${meta.handle}. Path: ${meta.path}`
    )
  }
  if (meta.kind === 'html') {
    return (
      `Opened ${meta.fileName} as an editable html session: ${String(meta.lineCount)} lines, ` +
      `${String(meta.headingCount)} heading(s), ${String(meta.linkCount)} link(s)` +
      `${meta.title ? `, title "${meta.title}"` : ''}. ` +
      `${meta.warnings[0] ?? ''}Handle: ${meta.handle}. Path: ${meta.path}`
    )
  }
  return (
    `Opened ${meta.fileName}: ${meta.blockCount} blocks, ${meta.wordCount} words ` +
    `(${meta.charCount} characters). Handle: ${meta.handle}. Path: ${meta.path}`
  )
}
