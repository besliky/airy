# Airy Copilot — MCP server for coding agents

`packages/mcp-server` (npm name `@airy-office/mcp`) is a Model Context
Protocol server that lets CLI coding agents work with real office files:
headless `.docx` / `.xlsx` / `.pptx` / Markdown / HTML editing and `.pdf`
text extraction through the suite's own engines,
plus a live bridge into the running Airy desktop app. It speaks MCP over stdio,
runs as a plain Node process (no Electron, no display), and needs no
installed app for the headless tools.

## Building

```bash
# from the repo root, after `npm install`
npm run build -w packages/mcp-server     # or: -w @airy-office/mcp
```

The entry point is `packages/mcp-server/dist/index.js` (esbuild bundle).
Quick smoke test:

```bash
node packages/mcp-server/dist/index.js   # speaks JSON-RPC on stdio
```

Spreadsheet support (`.xlsx` / `.xlsm` / `.xls` / `.ods`) spawns the Rust
xlsx sidecar. The server finds it automatically in a repo checkout
(`apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar`, built by
`npm run build -w @airy-office/sheets`, which needs `cargo` on PATH) or in an
installed app; `AIRY_XLSX_SIDECAR` points at an explicit binary.

Legacy text formats (`.doc` / `.odt`) are converted through LibreOffice:
install `soffice` (e.g. `sudo apt install libreoffice`) or point
`AIRY_SOFFICE` at the binary. Without it, `.doc` still opens read-only as
extracted text; `.odt` fails with an actionable error.

## Connecting an agent

### Claude Code

Project scope (recommended; shared via the repo):

```bash
claude mcp add --scope project --transport stdio airy \
  -- node /abs/path/to/airy/packages/mcp-server/dist/index.js
```

Equivalent `.mcp.json` in the project root:

```json
{
  "mcpServers": {
    "airy": { "command": "node", "args": ["/abs/path/to/airy/packages/mcp-server/dist/index.js"] }
  }
}
```

User scope: the same command with `--scope user`. Verify with
`claude mcp list` (should report `connected`) and `/mcp`; tools are named
`mcp__airy__*`. A permissions preset that auto-approves reads and asks
before writes:

```json
{
  "permissions": {
    "allow": ["mcp__airy__read_*", "mcp__airy__open_*"],
    "ask": ["mcp__airy__*write*", "mcp__airy__live_*"]
  }
}
```

### ZCode

User scope — `~/.zcode/cli/config.json`:

```json
{
  "mcp": {
    "servers": {
      "airy": {
        "type": "stdio",
        "command": "/abs/path/to/node",
        "args": ["/abs/path/to/airy/packages/mcp-server/dist/index.js"],
        "enabled": true,
        "timeoutMs": 60000
      }
    }
  }
}
```

Workspace scope is the same JSON in `<repo>/.zcode/config.json`. If Node
comes from nvm, give its absolute path in `command` (a bare `node` may not
resolve). Verify under Settings → MCP, or check
`grep 'mcp.tools.registered' ~/.zcode/cli/log/zcode-$(date +%F).jsonl`.

### Other agents

Any MCP client that supports stdio servers works: server name `airy`
(letters/digits/dash only), no required environment variables, startup
under 5 s, logs on stderr, reads carry `readOnlyHint`.

## Run the MCP server from an installed Airy app

Every installer — Windows, Linux and macOS — bundles the built server as
`resources/mcp/index.js` next to the app binary (on mac:
`Airy.app/Contents/Resources/mcp/index.js`), and the app itself doubles
as its Node runtime: with `ELECTRON_RUN_AS_NODE=1`, `Airy.exe` / `airy` /
`Airy.app/Contents/MacOS/Airy`
runs any script exactly like `node` (same Node line the bundle targets, and
`process.resourcesPath` still points at the install's `resources` dir, so
the xlsx sidecar in `resources/native` is found automatically). A machine
with the app installed therefore needs no Node.js and no repo checkout.
The bundle is self-contained — nothing else from the install is required.

**Windows (NSIS).** The installer is per-user (assisted mode with a
changeable directory), so the default location is
`C:\Users\<you>\AppData\Local\Programs\Airy`. Claude Code `.mcp.json`:

```json
{
  "mcpServers": {
    "airy": {
      "command": "C:\\Users\\<you>\\AppData\\Local\\Programs\\Airy\\Airy.exe",
      "env": { "ELECTRON_RUN_AS_NODE": "1" },
      "args": ["C:\\Users\\<you>\\AppData\\Local\\Programs\\Airy\\resources\\mcp\\index.js"]
    }
  }
}
```

**macOS (dmg).** Drag the app to `/Applications`; the binary is
`/Applications/Airy.app/Contents/MacOS/Airy`, the server
`/Applications/Airy.app/Contents/Resources/mcp/index.js`. ZCode
(`~/.zcode/cli/config.json`):

```json
{
  "mcp": {
    "servers": {
      "airy": {
        "type": "stdio",
        "command": "/Applications/Airy.app/Contents/MacOS/Airy",
        "env": { "ELECTRON_RUN_AS_NODE": "1" },
        "args": ["/Applications/Airy.app/Contents/Resources/mcp/index.js"],
        "enabled": true,
        "timeoutMs": 60000
      }
    }
  }
}
```

**Linux (deb).** The package installs under `/opt/Airy`: binary
`/opt/Airy/airy`, server `/opt/Airy/resources/mcp/index.js`. ZCode
(`~/.zcode/cli/config.json`):

```json
{
  "mcp": {
    "servers": {
      "airy": {
        "type": "stdio",
        "command": "/opt/Airy/airy",
        "env": { "ELECTRON_RUN_AS_NODE": "1" },
        "args": ["/opt/Airy/resources/mcp/index.js"],
        "enabled": true,
        "timeoutMs": 60000
      }
    }
  }
}
```

Claude Code on Linux takes the same paths in `command` / `args` with the
same `env` (see the Windows snippet). The `rpm` installs to the same
`/opt/Airy` layout.

**Linux (AppImage).** The squashfs mount point differs between runs, so a
config cannot point into a mounted AppImage reliably. Either install the
deb, use a checkout, or extract the image once and point at the extraction:

```bash
chmod +x Airy-<version>-x86_64.AppImage
./Airy-<version>-x86_64.AppImage --appimage-extract
# runtime: ./squashfs-root/airy
# server:  ./squashfs-root/resources/mcp/index.js
```

The `live_*` tools still need the Airy app running; it bridges regardless
of which copy of the server connects to it.

## Tools

Headless (no app required):

| Tool                 | Signature (short)                                                                                                                                                                                                            |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ping`               | `()` — liveness probe                                                                                                                                                                                                        |
| `open_document`      | `(path)` — open `.docx/.xlsx/.xlsm/.xls/.ods/.doc/.odt/.md/.markdown/.html/.htm/.pptx/.pdf`, returns a session handle + meta                                                                                                 |
| `read_document`      | `(handle, blocks? \| range?)` — block overview or full restricted HTML for text documents; markdown: heading structure; html: parse5 structure summary (lines)                                                               |
| `read_workbook`      | `(handle, sheet?, range?)` — sheet overview or a pipe table of an A1 range                                                                                                                                                   |
| `read_deck`          | `(handle, slide?)` — deck overview (`index\|elements\|preview` per slide) or one slide's element list + full text                                                                                                            |
| `insert_content`     | `(handle, html, at?)` — docx: restricted-HTML fragment after block `at`; markdown: `text` + `at`/`afterHeading`/`marker`; html: verbatim fragment at `at`/`marker`; slides: `text` + `slide` (+ `slideElement`/box geometry) |
| `apply_ops`          | `(handle, ops, dryRun?)` — validated, atomic batch of edit ops (max 100): block ops for docx, line ops for markdown/html                                                                                                     |
| `apply_workbook_ops` | `(handle, edits, dryRun?)` — validated batch of cell edits: value / formula / style / rich text (max 100)                                                                                                                    |
| `save_document`      | `(handle, path?, overwrite?, format?)` — atomic save; refuses existing targets without `overwrite: true`; `format: "origin"` exports back to the legacy format                                                               |
| `close_document`     | `(handle)` — close the session, clean up temp files                                                                                                                                                                          |

Live (app running; always target the _active_ tab):

| Tool               | Signature (short)                                                    |
| ------------------ | -------------------------------------------------------------------- |
| `live_status`      | `()` — bridge ping, protocol version, open documents                 |
| `live_get_context` | `()` — active document's blocks, selection (`<sel>`), comments, path |
| `live_apply_ops`   | `(ops?, html?)` — one combined edit turn (html first, then ops)      |
| `live_undo`        | `()` — revert the last agent turn                                    |

In `live_apply_ops` the html is inserted at the **end** of the document, so
block indexes from `live_get_context` stay valid when the ops run — the ops
address the pre-insert numbering.

Edit ops are flat records targeting blocks by `nodeType` / `headingLevel` /
`containsText` / `blockIndexes`; `apply_ops` output includes the full
signature list. Block indexes shift after inserts — re-read before further
addressing (the live html insert is the carve-out: it appends at the end, so
earlier indexes are unaffected). Both `nodeType` vocabularies are accepted
everywhere, headless and live: the ops-guide names (`heading` / `paragraph` /
`listItem` / `image`) and the renderer's canonical names (`docHeading` /
`docParagraph` / `docListItem` / `image`) target the same blocks.

The live bridge additionally accepts the embedded registry's extra ops —
`setImageProperties` (resize/align image blocks) and `insertToc` (insert a
TOC field after a block) — and `setFont` / `setMatchedFont` gain a
`link: { url } | null` field. Live `apply_ops` / `insert_content` only work
on the ACTIVE tab when it is a **docs** document; a sheets/slides/pdf tab in
front answers `not_docs_tab` (use the headless workbook tools for
spreadsheets).

Workbook editing goes through `apply_workbook_ops`: one batch of cell edits,
each targeting a single cell by sheet (name or index) and A1 ref with a
value, a formula (stored without a cached result, so apps recalculate on
open), a style patch (`bold`, `fillColor`, `numberFormat`, borders, ...),
rich-text runs, or a combination. Formulas win over values; later edits to
the same cell win per channel (content replaces content, style replaces
style). Charts, pivots, merged ranges and sheet structure are **not**
editable headlessly.

Markdown editing (`.md` / `.markdown`) is line-based. `read_document` shows
stats (including the file's EOL style and BOM), the heading list
(`ordinal|line|level|text` — ATX headings outside fenced code blocks and
YAML front matter, setext `===`/`---` headings are plain lines; capped at
200 entries with 80-character texts) and the full text; the whole read
shares the 30k budget (`blocks`/`range` select lines). `insert_content` takes markdown `text`
(not `html`) at one of three positions — after the first line containing
`marker`, after heading N (`afterHeading`, 1-based ordinal from the read),
or after line `at` (`-1` = start; default: end). Appending at the end
preserves the file's trailing-newline shape: a file that ends with a newline
keeps it (no blank line is added), a file without one keeps ending without
one, and an empty file gains no leading newline. `apply_ops` runs line ops
instead of block ops: `insertLines after text`, `replaceLines from to text`
(empty text deletes the range), `deleteLines from to`, and `findReplace find
replace matchCase? from? to?` (line-scoped — find and replace must be
single-line; optional inclusive line window). Line counts and indexes cover
real lines (a final newline does not create an extra empty final line);
they are 0-based and shift after every splice — re-read between
edits. Encoding: UTF-8 only (BOM-prefixed UTF-8/UTF-16 opens; files that are
not valid UTF-8 are refused with a conversion hint), a leading BOM survives
saves, untouched lines keep their exact bytes (EOLs included — a CRLF file
stays CRLF, mixed line endings keep their own), and a zero-edit save writes
the original bytes back verbatim. Markdown files cap at 8 MiB and
2,000,000 lines (the size is checked via stat before the file is read, the
line count before the line model is built).

HTML editing (`.html` / `.htm`) is line-based, like a source editor.
`read_document` shows the title, stats (including the file's EOL style and
BOM), a parse5 structure summary — headings (`ordinal|line|level|text`) and
links (`ordinal|line|text -> href`) with 0-based line positions, from the
same parser the Airy HTML editor builds on — and the full text; heading and
link lists cap at 200 entries each and share the 30k read budget with the
text (`blocks`/`range` select lines). `insert_content` splices the fragment
**verbatim** (no reparse or rewrite — exactly what you send lands on disk,
modulo the file's EOL style) after the first line containing `marker` (e.g.
`</body>` to append rendered content) or after line `at` (`-1` = start;
default: end). `afterHeading` is a markdown-session option — passing it to an
html session is an explicit error (html has no heading addressing; position
via `marker` or `at`). Appending at the end preserves the file's
trailing-newline shape (no blank line added, an empty file gains no leading
newline), and line counts and indexes cover real lines — a final newline
does not create an extra empty final line. `apply_ops` runs line ops
instead of block ops: `insertLines after text`, `replaceLines from to text`
(empty text deletes the range), `deleteLines from to`, and `findReplace find
replace matchCase? from? to?` (line-scoped — find and replace must be
single-line; optional inclusive line window). Line indexes are
0-based and shift after every splice — re-read between edits. Encoding:
UTF-8 (BOM-prefixed UTF-8/UTF-16 accepted); bytes that are not valid UTF-8
open only when the document declares a usable `<meta charset>` — undeclared
non-UTF-8 is refused with a conversion hint. A leading BOM survives saves,
untouched lines keep their exact bytes (EOLs included — a CRLF file stays
CRLF, mixed line endings keep their own), and a zero-edit save writes the
original bytes back verbatim; an edited save of a legacy-charset original
writes UTF-8 **and rewrites the charset declaration to `utf-8`** (with a
warning) — browsers trust the declaration, so leaving a stale legacy claim
would render the saved file as mojibake. HTML files cap at 8 MiB and
2,000,000 lines (checked via stat / a counting pass before anything is
materialized); documents above 1M characters skip the structure scan (read
shows the text only).

Slides editing (`.pptx`) goes through the suite's pptx engine. `read_deck`
without options returns the deck overview — `index|elements|content preview`
per slide plus deck stats; with `slide` (0-based index) it returns that
slide's element list (`index|type|name|text preview`; types `text`, `shape`,
`picture`, `group(n)`, `table(rxc)`, `chart`, `passthrough`, placeholders as
`text:title` / `shape:body`) and the slide's full text (groups and table
cells included). `insert_content` takes plain `text` plus `slide`: with
`slideElement` the text replaces that element's body (text boxes and
autoshapes; line breaks become paragraphs, a first text on a bare autoshape
gets PowerPoint's centered authoring defaults, connectors refuse — they
cannot hold text), without it a new text box is added at `x`/`y`/`width`/
`height` inches (default 6 x 1 in at 1", 1"). `apply_ops` is not available
for slides (the rich op registry is app-side); pictures, tables, charts,
groups and slide structure (add/remove/reorder slides) are not editable
headlessly. Legacy `.ppt`/`.odp` are refused with a conversion hint
(`soffice --convert-to pptx`). Saves keep untouched zip entries
byte-identical and a zero-edit save writes the original bytes back verbatim.

PDF reading (`.pdf`) is extraction-only: the document opens read-only
(`editable: false`) with text extracted by pdfjs through the file-parse
package (pages separated by blank lines; scanned/image-only pages extract no
text). `read_document` shows the text, saving is refused — headless PDF
editing is out of scope.

## Live mode

When the Airy app runs, its main process starts a bridge server on a local
socket (UDS on macOS/Linux, named pipe on Windows) and writes
**`airy-bridge.json`** into the app's userData directory:
`{socketPath, token, pid, protocolVersion}`, file mode `0600`, a fresh
random token per app session. The MCP server discovers the file by trying
`<appData>/Airy/airy-bridge.json`, then `Airy Dev`, then the legacy
`GenOffice` / `GenOffice Dev` layouts; `AIRY_BRIDGE_FILE` overrides the
location exclusively (no fallback — set it when the app's userData is
redirected, e.g. `AIRY_USER_DATA` in dev).

The client sends the token with every call and rereads the info file on
each connect, so an app restart (new token) never authorizes a stale
connection. Per-call timeout is 30 s. Bridge messages cap at 8 MiB per NDJSON
line in both directions: a request over the cap is rejected and the connection
closed, while a result that would exceed it (for example a `live_get_context`
of a very large document) comes back as an `invalid_request` saying the
response is too large — the connection stays usable, so narrow the request
(fewer blocks, smaller ranges) and retry. Live edits are visible immediately;
with track changes on they are authored as "Airy Copilot", and each bridge
call is one undo step (`live_undo` after a combined `html` + `ops` call
needs two undos). A combined `live_apply_ops` whose ops batch fails after
the html was inserted rolls the insert back with an automatic undo — the
document ends at its pre-call state (or the error says the edit may be
partially applied and to call `live_undo`). A `stale_document` error means
the user edited the document since your last `live_get_context` — fetch
fresh context. Bridge turns are attributed to the connection that made
them: the automatic rollback after a failed combined `live_apply_ops` only
reverts that client's own turn — when another copilot client edited in
between it refuses (`turn_owned_by_other`) and the error says the insert
remains. An explicit `live_undo` may still revert another client's turn (a
deliberate choice) and its result says whose turn it was (`anotherClient`).

`AIRY_DISABLE_BRIDGE=1` turns the bridge off in the app entirely.

## Environment variables

| Variable              | Side   | Meaning                                                               |
| --------------------- | ------ | --------------------------------------------------------------------- |
| `AIRY_WORKSPACE_ROOT` | server | Confinement root for all input/output paths (default: process cwd)    |
| `AIRY_BRIDGE_FILE`    | server | Explicit path to `airy-bridge.json` (disables default discovery)      |
| `AIRY_XLSX_SIDECAR`   | server | Explicit path to the `xlsx-sidecar` binary                            |
| `AIRY_SOFFICE`        | server | Explicit path to the LibreOffice `soffice` binary                     |
| `AIRY_DISABLE_BRIDGE` | app    | `=1` disables the live bridge server in the desktop app               |
| `AIRY_USER_DATA`      | app    | Redirects the app's userData (dev/test); pair with `AIRY_BRIDGE_FILE` |

## Formats and limitations

| Format              | Open                                                             | Save                                                                                                                                                  |
| ------------------- | ---------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `.docx`             | native, fully editable                                           | byte-preserving `.docx`                                                                                                                               |
| `.xlsx` / `.xlsm`   | native (Rust sidecar), formulas recalc                           | `.xlsx`                                                                                                                                               |
| `.xls` / `.ods`     | converted import (calamine) — **styles are lost**, tools warn    | `.xlsx` sibling; `format: "origin"` best-effort `.ods` via soffice; true `.xls` output is not supported                                               |
| `.doc`              | via soffice → editable `.docx`; without soffice → read-only text | `.docx`; `format: "origin"` best-effort `.doc` via soffice                                                                                            |
| `.odt`              | via soffice → editable `.docx` (without soffice: clear error)    | `.docx`; `format: "origin"` best-effort `.odt` via soffice                                                                                            |
| `.md` / `.markdown` | native (UTF-8, BOM accepted; invalid UTF-8 refused)              | line-preserving UTF-8; zero-edit saves round-trip verbatim; UTF-16 originals convert to UTF-8 on save                                                 |
| `.html` / `.htm`    | native (UTF-8, BOM accepted; declared legacy charsets accepted)  | line-preserving UTF-8; zero-edit saves round-trip verbatim; legacy/UTF-16 originals convert to UTF-8 (legacy charset declarations rewritten to utf-8) |
| `.pptx`             | native (pptx-engine; text boxes and shape text editable)         | byte-preserving `.pptx`; zero-edit saves round-trip verbatim                                                                                          |
| `.pdf`              | read-only text extraction (pdfjs via file-parse)                 | not editable (saving is refused)                                                                                                                      |

Byte preservation differs by format. `docx` saves keep untouched parts
byte-identical, and a zero-edit save writes the original bytes back verbatim.
`markdown` and `html` sessions behave the same at line granularity: untouched lines keep
their exact bytes (EOLs included) and a zero-edit save round-trips the file
verbatim; an edited save writes UTF-8 with the original BOM re-applied (html
sessions also rewrite a legacy charset declaration to `utf-8`, so the saved
file renders correctly in browsers).
`pptx` saves keep untouched zip entries byte-identical (dirty slides are
rebuilt from their byte anchors) and a zero-edit save writes the original
bytes back verbatim.
`xlsx` saves keep untouched zip entries byte-identical **except
`xl/workbook.xml`**: the save gateway always ensures the `fullCalcOnLoad`
flag so edited formulas recalculate on open, so even a zero-edit workbook
save may rewrite that one entry — and the save result's `unchanged` flag is
journal-based for workbooks (no edits journaled), not a byte guarantee.

Reads are bounded to keep tool answers inside the ~30k-character MCP budget:
`read_document` truncates its output at 30,000 characters (markdown/html
reads count the structure summary toward that budget — heading and link
lists cap at 200 entries each; the block overview tightens previews and
elides the middle first; a selected-blocks read tells you to narrow the
range), `read_document`'s `blocks` parameter accepts at most 200 indexes per
call, a `range` may span at most 10,000 blocks/lines (a larger span is
rejected up front — split it into several reads), and `read_workbook` ranges
cap at 20,000 cells (split larger ranges into smaller reads). `read_deck`
answers share the same 30k budget (the deck overview tightens previews and
elides the middle first; a slide detail read truncates), and slides
`insert_content` text caps at 200,000 characters. Markdown and
HTML sessions add an 8 MiB / 2,000,000-line open cap (larger files are
refused with a clear error, by stat before the content is read); HTML
documents above 1M characters skip the parse5 structure scan.

## Security model

- **Path confinement.** Every input and output path must resolve inside the
  workspace root (`AIRY_WORKSPACE_ROOT`, default the server's cwd); traversal
  that escapes the root is rejected with a clear error. Each session pins the
  root at open time and keeps confining its saves against that root, so a
  later `AIRY_WORKSPACE_ROOT`/cwd change never re-confines a live session.
  When that pinned root itself disappears mid-session (the workspace
  directory was moved or renamed), saves are refused with a stale-root error
  naming the root instead of silently re-creating the dead directory — for
  every session kind alike (docx, workbook, slides, markdown/html) — reopen
  the document from its new location and re-apply your edits.
  Symlinks are resolved
  for both the root and the candidate before the check, so a link that lives
  inside the root but points outside cannot smuggle paths out (links that
  resolve back inside the root stay usable). On Windows the comparison folds
  case, matching the case-insensitive filesystem — `c:\users\...` and
  `C:\Users\...` are the same path. Confinement is checked at resolution
  time: a racing local attacker with write access inside the root (swapping a
  checked directory for a symlink before the write lands) is out of scope.
- **Token, not location.** The bridge socket and its info file are `0600`;
  every bridge call must carry the current per-session token, which is
  reread from disk on every connect. The bridge listens on a local
  socket/named pipe only — no network surface.
- **No silent overwrites.** `save_document` is atomic (temp + promote) and
  double-fenced: saving over the opened file refuses with an error when the
  file changed on disk since it was opened (an external writer — another
  editor, sync client, or the Airy app itself), and an explicit save-as to a
  path that already exists is refused unless it is a file the session itself
  opened or saved — pass `overwrite: true` to replace an unrelated file. A
  guarded fresh target is promoted with an exclusive link, so a file that
  appears between the existence check and the write cannot be silently
  replaced either. The
  remedy for a fence error is to reopen and reapply, or to use the `live_*`
  tools when the document is open in the app.
- **Read-only until save.** Opening and editing never touch the original
  file; converted imports write a new sibling file and leave the original
  untouched.
