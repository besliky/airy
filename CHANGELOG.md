# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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

### Security

- Screen capture without a user-visible prompt is denied, and renderer
  file reads are confined to the document's own paths.
- API keys are encrypted at rest and masked in the renderer.
- Spawned Codex CLI paths are restricted to an allowlist.
- MCP workspace confinement resolves symlinks and Windows case
  differences before checking boundaries.

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

[Unreleased]: https://github.com/besliky/airy/compare/v0.9.3...HEAD
[0.9.3]: https://github.com/besliky/airy/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/besliky/airy/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/besliky/airy/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/besliky/airy/releases/tag/v0.9.0
