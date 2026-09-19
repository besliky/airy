# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.12.0] - 2026-09-19

### Added

- Docs: Compare documents as tracked changes (legal blackline) with run-level
  diffing, budgeted matching, and read-only gating; the differences pane remains.
- Docs: cross-reference types (text / page number / heading number) with F9 cache
  refresh and robust SEQ caption labels.
- Docs: section line numbers with Word restart semantics, paper-ink canvas overlay,
  and print parity.
- Docs: TOC options dialog (levels, page numbers, styles, hyperlinks) and Table of
  Figures authoring from SEQ captions.
- Docs: full styles gallery (Heading 1–9 + custom document styles) with Modify Style.
- Docs: footnote/endnote options (numbering formats, custom marks, conversion,
  navigation), hyphenation authoring, and the unequal-columns dialog with separators.
- Docs: area, scatter, bubble, and doughnut charts with chart edit mode and embedded
  workbook sync; alt text for pictures/shapes/tables; picture/shape shadow effects
  and borders; Compress Pictures.
- Docs: SmartArt insert (four presets) as real OOXML diagram parts.
- Sheets: range moves on imported workbooks (2D rectangle moves with full reference
  remapping and fail-closed guards) completing column/row move parity.
- Pagination: a grid compatibility profile — Word parity lifted to 95.9% page-start
  matches (page counts 25/25) and a LibreOffice profile reaching 86.0%.
- MCP: headless markdown and HTML sessions land in 0.11.0 tooling; this release
  hardens them (size/line caps, pinned workspace roots, deterministic fixtures).

### Fixed

- Bridge/MCP: requests are no longer dispatched after a connection closes; dribbled
  chunks coalesce without quadratic copying; oversized answers return a typed error
  instead of dropping the connection.
- Docs: wildcard search hardened (no catastrophic backtracking, placeholder nodes
  untouched); compare refuses to run over pending revisions; read-only documents are
  no longer mutated by anchors/links/compare.
- Sheets: whole-column refs in conditional formatting/validation/protected ranges
  remap on moves; structural replay remaps sortState; move ops are bounded by the
  sheet edges.
- UX: WCAG-compliant focus rings in sheets/slides; modal dialog semantics (focus
  trap, Escape) across new dialogs; busy states for compare and PDF page inserts;
  split-button and toggle screen-reader announcements.

### Changed

- Sheets renderer bundle split into layer chunks (largest 3.4 MB, eager −12% after
  lazy locale dictionaries); docs renderer entry −43% with lazy locale dictionaries.
- html2docx tests run browser domains in parallel; e2e waits are deterministic
  (44 fixed sleeps → 6 justified); byte-deterministic MCP fixtures.

## [0.11.0] - 2026-09-18

### Added

- Docs: Find & Replace wildcards (Word `*?[]` syntax, lazy semantics) and an
  ignore-diacritics option.
- Docs: internal hyperlinks — Insert → Link gains a "Place in This Document" tab
  (headings + bookmarks), ⌘/Ctrl+click jumps to the target.
- Docs: Sort — tables and selected paragraphs (multi-level, text/number/date with
  Word parsing conventions, header-row option).
- Docs: Insert → SmartArt — four presets (block list, vertical bullet list, basic
  process, organization chart) written as real OOXML diagram parts with live
  previews and node editing.
- Sheets: column moves on imported workbooks (drag a column header; references,
  merges, CF/DV, names, chart series, anchors and tables remap on save; undoable).
- Shell: live file search on the Home screen (⌘/Ctrl+F, localized empty state).
- PDF: Insert Pages from Another PDF with a preview dialog, position and page-range
  selection.
- MCP: headless markdown and HTML sessions (open/read/insert/apply_ops/save with
  EOL/BOM preservation and structure summaries; parse5 for HTML).
- UI: reusable modal dialog semantics (role=dialog, focus trap, Escape, focus
  return) adopted by the docs insert/sort/link and pdf insert-pages dialogs;
  FindPanel toggles announce state via aria-pressed.

### Fixed

- Live bridge: NDJSON frames decode across UTF-8 chunk boundaries (large CJK
  payloads no longer corrupt into U+FFFD).
- Slides/pptx: .pptx saves are atomic (temp file + rename) — a crash or full disk
  can no longer truncate the deck.
- Sheets: CSV save-back is atomic; whole-column/row refs in CF/DV/allow-edit
  ranges remap on moves; the column fast path no longer skips rows whose cell
  text contains ` r="`.
- Docs: wildcard search no longer matches inline-object placeholders (Replace can
  no longer destroy hard breaks/images/math), no longer backtracks
  catastrophically (linear matcher), and heading anchors are no longer stamped
  into read-only documents.
- MCP: read ranges are validated before allocation (huge `end` values fail fast);
  legacy-charset HTML saves rewrite the declaration to UTF-8; heading lists are
  capped and counted into the 30k budget; `afterHeading` is rejected on HTML
  sessions.
- Exports: html export filters link hrefs through the scheme policy (stored
  `javascript:` hrefs can no longer execute from exported files); pdf/markdown/
  html/docs export writes are atomic.

### Changed

- html2docx test suite runs browser domains in parallel (147s → 37s; full serial
  `npm test` 260s → 138s).

## [0.10.0] - 2026-09-17

### Added

- Docs: navigation-pane search — type to live-filter headings with matched-text
  emphasis and a distinct no-match state.
- Slides: threaded comments — replies nest under their parent with a
  per-thread composer, and threads can be resolved (muted + collapsed) and
  reopened; the ribbon badge counts threads.
- Slides: richer find — per-hit navigation with whole-word matching (same
  Unicode word definition as docs), find-previous, and a visible highlight
  of the matched range.
- Slides: File > Export as PDF now produces vector PDFs with real,
  selectable, searchable text instead of page bitmaps.
- Sheets: a print dialog and a File tab in the ribbon.
- Sheets/Slides: Help menu entries with a searchable keyboard-shortcuts
  dialog and Online Documentation / Copilot Guide links (localized in all
  20 locales).
- Shell: middle-click closes a tab, tabs get a context menu and
  reordering/next/previous accelerators.
- Shell: the previous tab session is restored on launch and window bounds
  are remembered across launches.
- Shell: periodic crash-recovery copies for Markdown and HTML tabs (every
  30 s and on blur while dirty), with the same Restore/Discard prompt as
  docs.
- Shell: styled error toasts instead of raw `alert()` dialogs.
- Shell: a Settings → General toggle for the copilot live bridge (on by
  default, replacing the env-var-only opt-out).
- Shell: configurable author name for comments and revisions.
- MCP: `apply_workbook_ops` tool for headless `.xlsx` editing (validated
  batches of cell edits with `dryRun` support).
- Accessibility: ARIA ribbon structure with keyboard navigation in sheets,
  and right-to-left chrome for Arabic and Hebrew.
- Localization: PDF fill-form strings and remaining sheets screen-reader
  labels translated for all 20 locales.
- macOS builds (arm64 `.dmg` and `.zip`) published by the release workflow,
  unsigned — see the README for the Gatekeeper caveat.
- README: real screenshots of the home screen and each editor, generated
  from an e2e spec and committed under `docs/assets/screenshots/`.

### Changed

- PDF→Word conversion runs in an Electron utility process, so the app
  stays responsive during long conversions.
- Recent files are validated through a short-lived asynchronous cache
  instead of synchronous `stat` calls, and workbooks are delivered as soon
  as the sheets editor reports its menu ready (replacing a fixed retry
  nudge).
- Deleting rows/columns referenced by formulas now succeeds and rewrites
  the formulas to `#REF!`, matching Excel semantics (previously refused).
- MCP `save_as` refuses to overwrite an existing file and rolls back
  partial live edits on failure.

### Fixed

- Tabs survive renderer crashes and main-process errors instead of taking
  the window down.
- Untitled sheets and PDF files are no longer written into the save
  directory on close/quit.
- Close-tab and quit accelerators are consistent across platforms.
- Slides enforce a minimum window size, preventing an unusable tiny state.
- Tertiary text and muted icons meet WCAG contrast.
- The README updater section now matches the shipped notify-before-install
  updater behavior.
- macOS installers bundle the MCP server, so the copilot CLI works from
  installed macOS builds.
- Ctrl+1..9 switch tabs from inside editor views, not just from the Home
  tab.
- Sheets: Cmd/Ctrl+Y redoes, matching Excel.
- Sheets: cross-sheet conditional-formatting and data-validation formulas
  are rewritten on row/column deletion too, not just cell formulas.
- Slides: SVG export mirrors flipped pictures, tables, charts, and groups.
- Slides: a single 360° pie wedge renders as a closed full circle instead
  of disappearing.
- Slides: find highlights project correctly through rotated and flipped
  ancestor groups.
- File menus lead with the New submenu and the macOS dock menu gains
  New HTML.
- Save-as works on volumes without hardlink support (exFAT, FAT, some
  network shares) instead of failing with an opaque error.

### Security

- Screen capture without a user-visible prompt is denied, and renderer
  file reads are confined to the document's own paths.
- API keys are encrypted at rest and masked in the renderer.
- Spawned Codex CLI paths are restricted to an allowlist.
- MCP workspace confinement resolves symlinks and Windows case
  differences before checking boundaries.
- Inserted links pass a scheme whitelist (http/https/mailto, `#fragments`,
  relative refs) on every AI/bridge/ops insert path; disallowed schemes
  like `javascript:` degrade to plain text.
- Mutating and privilege-granting IPC channels validate their sender
  (`home:*` answers the Home tab only; `win:new` and per-app recents answer
  that app's renderers only), and attachment folder grants require a
  witnessed user drop or paste.

## [0.9.3] - 2026-09-15

### Fixed

- Legacy-charset text files (e.g. GBK, Shift-JIS) are decoded correctly
  instead of producing mojibake or parse failures.
- A leading BOM is preserved as `U+FEFF`, matching `readFile(utf8)`
  behavior, when reading text files.
- Development: `@airy-office/ui` subpath imports resolve under vitest.

## [0.9.2] - 2026-09-11

### Changed

- The real Airy mark is used across the UI, and app icons were
  regenerated from the icon-only mark.

## [0.9.1] - 2026-09-10

### Added

- The MCP server is bundled into every installer (`resources/mcp`), so
  coding agents can run it from the installed app without Node.js.

### Changed

- Rebrand to Airy: workspace packages renamed from `@genoffice/*` to
  `@airy-office/*`, user-facing strings rebranded, the Genspark badge
  replaced by the Airy mark, application icons from brand assets, and a
  Download section added to the README.

## [0.9.0] - 2026-09-10

Initial fork release. Airy is an independent, locally-focused fork of
[GenOffice](https://github.com/genspark-ai/genoffice) centered on
agent-driven document work.

### Added

- MCP server (`@airy-office/mcp`, bin `airy-mcp`): a Model Context
  Protocol server over stdio giving coding agents headless access to the
  suite's engines — docx open/read/insert/apply_ops/save, xlsx via the
  Rust sidecar (read/save with recalc), and legacy/ODF import via
  conversion.
- Live bridge: the desktop app listens on a local socket (UDS / named
  pipe), publishes a token file, and exposes `live_*` MCP tools that edit
  the active document in the running app.
- In-app updater backed by GitHub Releases that only notifies; the
  download and the install each need explicit consent. Check-for-updates
  entry in the Help menu.
- Release workflow publishing Linux and Windows builds, plus a release
  guide.

### Removed

- The Genspark account login and cloud projects.
- The auto-updater and usage analytics from upstream.
- The `gsk` CLI search backend and the `@genspark/cli` dependency.
- The Genspark AI provider — AI in the app is bring-your-own-key only.

[Unreleased]: https://github.com/besliky/airy/compare/v0.12.0...HEAD
[0.12.0]: https://github.com/besliky/airy/compare/v0.11.0...v0.12.0
[0.11.0]: https://github.com/besliky/airy/compare/v0.10.0...v0.11.0
[0.10.0]: https://github.com/besliky/airy/compare/v0.9.3...v0.10.0
[0.9.3]: https://github.com/besliky/airy/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/besliky/airy/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/besliky/airy/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/besliky/airy/releases/tag/v0.9.0
