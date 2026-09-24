# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.19.0] - 2026-09-24

### Added

- Shell (Home): project catalogs stop silently truncating at 256 files — the
  `home:stat-paths` cap rises to a production level (`HOME_PATHS_CAP = 20,000`,
  one number shared by main and renderer in `shared/home-paths.ts`), the catalog
  loads through a chunked loader (500-path `statPaths` round-trips with a
  main-thread yield between chunks, so rows appear progressively and the window
  never blocks in one long task), and a catalog beyond the cap shows an honest
  localized "{n}+ files" counter (new `fileCountOver` key in all 20 locales)
  instead of a silently short total; client-side search now finds files beyond
  the old cap (a 300-file project indexes fully and "p0280" is found) (PR #179).

### Performance

- Slides: a shape drag on a 200-slide deck holds 60fps — frame p95 drops
  100 → 16.8 ms, gaps >33 ms fall 24–27 → 5–8 per gesture (the raster
  transitions at the gesture boundaries plus the drop commit; mid-drag frames
  hold 16.7 ms), drag wall 2.4 s → 0.97 s, RSS flat ~1.1 GB. The cost was
  per-dragmove full-layer redraws (O(painted pixels), ~80 ms/frame under
  software rasterization); four scoped measures cut it: snap guides/spacing
  state is value-compared so an unsnapped move re-renders nothing, the slide
  base rects are cached, layer rasters decimate to 0.3× while the drag is live
  and snap back at gesture end, and the selection Transformer hides for the
  gesture (PowerPoint-style); move history still commits exactly once per
  gesture, now pinned by 14 new unit tests (PR #183).

### Fixed

- Markdown:
  - Saving a .md file no longer silently destroys raw HTML the GFM schema has
    no node for: Ctrl+S without edits shrank a 298-byte file (div/script/style
    blocks, an inline script, a script inside a table cell) to 155 bytes, and
    `kbd`/`mark`/`details`/`span` degraded to bare text — the badges and
    `<details>` sections of real READMEs. A new `rawHtml` atom node with a
    low-priority catch-all parseDOM rule captures anything unclaimed, stores
    its `outerHTML` verbatim and writes it back on save; dedicated rules keep
    winning (`<br>` → hard break, `<img>` → image), chips render as escaped
    text so stored scripts stay inert, docx export keeps chips as literal
    source runs, and the segmented parse (#150) stays bit-identical to the
    monolithic one; reopen is a fixed point (PR #189).
  - Find/replace on a giant single-paragraph document stops freezing and
    killing the renderer (a 1 MB one-paragraph .md with 43,479 matches):
    Replace All crashed the renderer in 4/4 runs and typing in the find panel
    blocked >30 s — painted search hits are capped at 500 (active match, then
    the visible viewport, then the earliest) while the counter stays honest and
    every match stays navigable/replacable, the status bar shows a localized
    "+N more matches not highlighted" plus live Replace All progress (20
    locales), Replace All rebuilds each textblock as precomputed slab steps in
    ONE transaction (O(block) instead of O(matches × block), marks preserved,
    cancellable by closing the panel), and the search no longer builds a
    per-character position array. Renderer crashes 4/4 → 0/4; search max
    continuous main-thread block ~8.7 s → ~3.4 s (the residual giant-paragraph
    rebuild wall is tracked as PERF-1700) (PR #185).
- MCP & automation:
  - `read_document` on a docx session keeps its declared ~30k answer budget,
    which was breached 27–97× on any sizeable document: the default read's
    elide pass ran only once (2/3 of overview lines survived — an 857 KB answer
    for a 42k-block document) and the blocks/range path appended the FULL
    overview before the selected HTML (a five-block range read cost ~1.57 MB).
    The overview now goes through a repeated-elide `fitOverview` ladder (kept
    lines halve each round, one honest elided-count marker, block numbering
    verifiable at both ends), selection reads cap the overview at 4 KB and give
    the selected blocks' HTML the rest with an explicit truncation note,
    budgets are measured in UTF-8 bytes (char counting underestimated CJK
    threefold), and a final guard makes the budget unreachable-proof:
    857,103 → 19,783 bytes default, 1,574,163 → 2,850 bytes for a 5-block
    range; the response shape for small documents is unchanged (PR #186).
  - A renderer killed mid-AI-turn no longer leaves a zombie stream: the killed
    webContents survives with `isDestroyed() === false`, so the app-wide
    `ai:stream` handler kept draining the provider SSE to the end of the turn —
    minutes of paid tokens nobody receives, plus one "Render frame was
    disposed" error per delta. Every stream now registers under a per-sender
    registry (reusing the sheets BUG-407 `AiStreamRegistry`) that aborts that
    sender's whole registry on `render-process-gone` or `destroyed`; per-stream
    cancel semantics are unchanged and the graceful-teardown tests stay green
    (PR #188).
- Shell:
  - A click on an unreadable Home row (chmod 000, a missing path, or a
    directory wearing a document extension) reports the #154 localized error
    dialog instead of ending in total silence — no tab, no dialog, no toast:
    the Home open decision lives in a dependency-injected `home-open.ts`, and a
    failed open routes through the same classify/`reportOpenFailure` pair as
    the argv/second-instance paths (EACCES/EPERM/ENOENT/EISDIR, eisdir getting
    the "This is a folder" message); valid files open exactly as before
    (PR #180).
  - A failed forwarded open no longer cascades into ~21 identical "Could not
    open" dialogs: an unpacked launch whose single-instance lock request fails
    retries it 20×, and every `requestSingleInstanceLock` re-emitted
    `second-instance` with the same launchPath. Failure dialogs now dedupe per
    path while one is on screen (a different path keeps its own dialog;
    dismissing releases the path), and the retry loop no longer re-sends the
    launchPath additionalData — the first request already delivered it, so the
    source of the storm is gone, not just its echo (PR #181).
  - Zombie tabs self-heal: a kill -9'd docs renderer left a tab whose
    webContents is NOT destroyed, so `destroyed` never fired, the bridge list
    kept advertising the corpse (`active:true`), every bridge call into it
    burned the full 30 s timeout, and re-opening the file re-activated the
    corpse without a new webContents. Any renderer death under a live tab now
    marks the tab dead: it leaves the bridge list and every
    `find*TabByPath`, bridge calls fail in milliseconds with a typed
    `tab_closed` (30,024 ms → 0 ms live), and re-opening the file revives the
    document as a NEW tab with a fresh renderer (View > Reload also revives);
    verified live under xvfb (PR #184).
- Slides:
  - Vertical text in table cells renders rotated instead of horizontal:
    `a:tcPr vert="vert270"` (rotated column headers — a WYSIWYG break with
    PowerPoint on every such table), plus `vert`/`eaVert`/`wordArtVert`, was
    never mapped by the table parser while the whole vertical-layout render
    path already existed for shape text. `parseTableCell` now maps `tcPr@vert`
    onto the cell text body (tcPr wins over an unusual bodyPr@vert; unknown
    values stay horizontal), and the shared consumers apply it for free —
    Konva glyph rotation on canvas, rotate transforms in the SVG/PDF export,
    row auto-growth; save keeps the attribute byte-for-byte, and
    `wordArtVertRtl` stays intentionally unsupported (horizontal, as for shape
    text) (PR #182).
- Platform:
  - Atomic saves stop forking documents saved through a symlink path:
    `atomicWriteFile`/`renameDurably` wrote a temp next to the given path and
    `rename(2)`d it over — and rename does not follow symlinks, so saving
    through a link silently replaced the link with a regular file while the
    destination kept the old bytes: the document forked in two (confirmed live
    on the markdown editor via the staleness-fence Overwrite and on the html
    editor with no dialog at all). Both helpers now resolve a live symlink
    first and land temp+rename on the REAL target — the link survives and
    readers through it see the save. A dangling link is written through in
    place by `atomicWriteFile` (exactly like a plain write) and refused with
    the new `BrokenSymlinkTargetError` by `renameDurably` instead of silently
    destroying the link; out-of-root symlink escapes stay refused by the
    existing MCP `resolveConfined` gate (PR #187).

## [0.18.0] - 2026-09-24

### Added

- Markdown & HTML: a manually chosen text encoding is remembered per file and
  survives reopen — `encoding-memory.ts` in the markdown and html main processes
  (shared app-settings single-writer, LRU cap 200) behind path-fenced
  `markdown:set-encoding` / `html:set-encoding` IPC; on open the precedence is
  BOM > explicit pick > auto-detection/meta charset. The in-renderer
  "Reopen with encoding" picker remains a named follow-up (PR #167).
- Docs: a pulsing "Loading…" badge (role=status, `appDocLoading`, 20 locales)
  appears in the status bar next to the counters while the tail of a large docx
  is still loading in phases, so in-progress "Page 1 of 4, 832 words" is no
  longer mistaken for truncation (PR #169).

### Changed

- Test infrastructure: the unit serial wall drops 172.6 s → 151.5/156.8 s (two
  consecutive warm runs, under the ≤160 s goal) via a bulk vitest project
  (`isolate: false`, module-graph reuse in workers) with a `NEED_ISOLATION`
  quarantine grown 22 → 25 files (global Univer DI graph conflicts under CI
  sharding); the shortcut-registry regression (2.42 → 5.15 s) is fixed with
  `Promise.all` over its 20 lazy imports, and PERF_BASELINE is updated (PR #164).
- Test infrastructure: local `test:e2e:xvfb` works again — the host's Wayland
  environment (WAYLAND_DISPLAY/XDG_SESSION_TYPE/GDK_BACKEND) leaked into
  `electron.launch`, Electron 43's ozone autodetect picked wayland and hung in
  the mojo handshake before `main.js`, timing out all 68 e2e specs;
  `scripts/run-e2e-xvfb.sh` now unsets the variables under xvfb, pins
  `ELECTRON_OZONE_PLATFORM_HINT=x11` and warns if the environment re-leaks:
  68/68 timeouts → 62 passed / 6 failed in 3.1 min with zero timeouts (the
  remainder is environmental freetype/docs-visual drift, verified independent);
  CI scripts untouched (PR #176).
- Test infrastructure, test-only: two pin tests lock the markdown/html `closeTab`
  renderer-reclamation contract after the PERF-1657 check returned "not
  reproducible / covered by #141" — multi-open of 15 md files returns exactly to
  baseline 4 processes / 615 MB (PR #170); the flaky sidecar prewarm test is
  determinized (the wait_for predicate waits for `!prewarm_active` beside the
  resident insert, 30 s poll deadline) — production code untouched, assertion
  not weakened (PR #171).

### Performance

- Sheets: the first edit after opening a large workbook lands in 0.59–0.91 s
  (was 5.3–7.5 s; the audit's 100k×20 book: 13.8 s → ~1 s) — a resident formula
  model is prewarmed in the background right after open (size gate 1–64 MB
  compressed, cache-lock held for the whole build, purge-epoch guard against the
  close×prewarm race); books over 64 MB (the 26.4 GB incident class) are never
  prewarmed (PR #161).
- Sheets: importing a 500k-row CSV (28.5 MB, 500,001×8) drops from
  16.5/16.3/20.2 s to 9.2/9.7/9.3 s engine-level (median 9.3 s, goal ≤15 s) —
  JSZip DEFLATE (6.5–8.7 s, ~50% of the wall) leaves the import path: a minimal
  OPC zip writer on `node:zlib` plus a single-pass cell XML escape cut the
  conversion itself from ~10.5 s to 2.9–3.4 s; a JSZip fallback keeps browser
  builds working (PR #163).

### Fixed

- Slides:
  - LibreOffice Impress accepts app-saved pptx again: the streaming zip writer
    (`generateNodeStream` with `streamFiles: true`) emitted data descriptors
    (general-purpose bit 3, zeroed crc/sizes in local headers), which strict
    loaders reject while python-pptx tolerates them — the same node stream now
    runs with `streamFiles: false`, the bytes are identical to the proven
    `savePptx` path (hash-verified on 180 MB of media), an integration test
    converts four saved decks (including a media-heavy one) to PDF through a
    real soffice, and peak memory is unchanged (one compressed entry in
    transit, not the deck) (PR #174).
  - Copy Slide carries the slide's speaker notes: the transfer bundle includes
    the notesSlide part and its relationships (minus the slide back-ref and
    notesMaster rel), the receiving deck materializes a fresh part and
    re-parents the target's notesMaster; slides without notes leave no tails,
    and duplicateSlide is untouched (PR #177).
  - Video export stops littering the destination folder: a single-owner module
    sweeps orphaned `.<name>.<12hex>.tmp` files older than one hour (by mtime)
    at the next export's stream-begin — main-only and fire-and-forget, so a
    killed export no longer leaves debris the startup sweeper never covered
    (PR #160).
- PDF:
  - A corrupt `/Rotate` value (a name instead of a number, some producers) no
    longer fails every annotate-save with a bare "Save failed": `pageRotation()`
    is a safe getter (never throws, snaps to multiples of 90, resolves indirect
    refs) at all seven read sites, and broken rotations are repaired at load;
    output is verified through pdf-lib reload, pdf.js and poppler (PR #172).
  - XMP metadata (catalog `/Metadata`) survives saves with content edits — text
    edit/insert, image edit and annotation deletion used to drop the stream:
    `applySaveRequest` extracts it from the source bytes before the pdfium
    stages and re-registers it byte-for-byte after the final load; plain rewrite
    saves never lost XMP (PR #175).
- Docs:
  - A refused PDF export (the read-only `canPdfWrite` gate) settles the status
    bar in a localized `appExportPdfFailed` message instead of leaving
    "Exporting PDF…" forever (PR #165).
  - Opening a docx with corrupt XML shows a short localized refusal ("File is
    corrupted: document.xml, position 1:1", 20 locales) instead of the raw
    fast-xml-parser validator dump; the full text goes to the console (PR #168).
  - The two color caret buttons in the ribbon (highlight / font color) carry
    aria-labels and tooltips via i18n keys in all 20 locales, pinned by a
    contract test (PR #166).
- Shell:
  - Renaming a file in Home reaches every window, not just the focused one: a
    rename broadcast walks all tab managers with per-editor renamed hooks
    (docs/sheets/slides/markdown/html), so a second window no longer keeps the
    old title and save path — Ctrl+S from it can no longer write to the old
    path (PR #173).
- Sheets:
  - The sidecar failure reply echoes the `requestId` recovered from the raw wire
    line (best-effort, string form; non-JSON keeps the legacy empty id), so
    direct protocol clients no longer wait out a 120 s timeout on an unparseable
    request — the `invalid_json` failure code is preserved (PR #162).

## [0.17.0] - 2026-09-24

### Added

- Markdown & HTML: in-place saves are fenced against external changes — a per-view
  stamp (mtime + size) taken at open (and refreshed after every save and shell-side
  rename) is re-checked as late as possible before the atomic write; on mismatch a
  native dialog offers Save As / Overwrite / Cancel, and autosave (tick/blur) is
  rejected silently via an additive `auto` wire flag instead of popping a dialog
  every 30 s (PR #151).

### Changed

- Breaking (formula semantics — shift to Excel 365): the sheets engine upgrades
  ironcalc 0.7.1 → 0.8.3, so SUMPRODUCT, LET and dynamic arrays evaluate natively,
  including file-level `_xlfn.LET(...)`. Comparisons lift element-wise over ranges,
  `=A1:A3` spills, and non-CSE `SUM(IF(A1:A3>10,1,0))` returns Excel 365's 2 instead
  of 1; SUMPRODUCT follows Excel exactly (text counts as 0, mismatched dimensions →
  `#VALUE!`), and unknown names still give `#NAME?`. The lockfile gains a single new
  transitive package (regex-lite); LAMBDA/FILTER/SORT/UNIQUE/SEQUENCE and spill UI
  semantics are named follow-ups (PR #155).

### Performance

- Markdown: giant files hydrate in phases — the body is split at safe top-level block
  boundaries, the first segment mounts synchronously and the rest hydrates in the
  background with typing enabled throughout; save/export/print wait for hydration and
  the result stays bit-identical to a monolithic parse. 20k-paragraph TTI
  25.4–25.8 s → 0.87–1.04 s, settled 27.6–27.9 s → 5.3–5.4 s (PR #150).
- HTML: on a 3 MB minified single-line document, TTI 39.7 s → 2.56 s and source typing
  ~22 s per key → 43–75 ms — the instrumented preview copy is built in one O(N) pass
  (the old splice loop copied ~96 GB of strings), sid matching is bucketed so
  parse-map rebuilds stay linear instead of O(elements²), and rebuilds/preview pushes
  leave the typing path behind a stale-serve map cache with a giant-document
  auto-rebuild limit (PR #149).

### Fixed

- Sheets:
  - Legacy .xls conversion carries row heights (ROW ht/fUnsynced, DEFAULTROWHEIGHT →
    sheetFormatPr) and hidden rows, so forms with non-default header heights keep
    their shape; zip/BIFF5 output stays byte-identical (PR #145).
  - Printing honors the declared row height: the declaration is authoritative
    (text-derived height only clamps from below, wrapped rows keep their text boost),
    so a border-template sheet lays out 3 → 2 pages with plan == render and the
    orphan stub page is gone (PR #146).
  - Legacy .ods/.xls import works end to end again: both styles.xml writers emit
    `<cellStyles>` from a single shared source, the ironcalc styles panic on import
    is gone, and formulas of imported workbooks compute (PR #148).
  - Deleting a source sheet applies Excel's #REF! semantics: dependent formulas are
    rewritten (`SUM(#REF!)`) on both delete paths, a save/reload round-trip shows the
    #REF! error instead of silently emptying, and undo restores the sheet (PR #156).
  - .ods import translates a conservative ODF-formula subset to xlsx syntax (the
    `of:=` prefix, bracketed `[.A1]`/`[Data.B1]` references, `;` → `,`,
    `COM.MICROSOFT.*` aliases); unrecognized formulas pass through verbatim with
    their cached value pinned, and .xlsx/.xls output stays byte-identical (PR #157).
- Markdown & HTML:
  - html2docx survives dirty HTML: an in-page normalize step splits inline elements
    with block children into (clone, lifted blocks, clone) before classification, so
    documents with unclosed p/li no longer collapse everything after the first h1
    into one paragraph; valid documents are a no-op and `<a>` is deliberately
    excluded (transparent content model) (PR #152).
- Text import:
  - GBK/gb18030 sentences no longer decode as windows-1250 soup: byte-honest script
    coverage plus a "soup signature" withholds single-byte bonuses when the bytes
    prove two-byte CJK, and scoring treats the newer ICU's PUA output as replacement
    so the charset choice is ICU-version-independent (Thai cp874 and big5 TW fixed as
    a bonus); on a 30-vector blind comparator the fix picks correctly 26/30 vs 16/30
    on main, with zero regressions (PR #147).
- Slides:
  - svgBlip pictures render from the vector part in the author's own frame instead of
    stretching a 1×1 PNG fallback across it; the embedded raster is kept as a
    decode-failure fallback, preloaded by editor/master/audience views; figure fills
    stay on the raster embed and save fidelity is untouched (PR #153).
- Shell:
  - Opening an unreadable file (EACCES/EPERM/ENOENT/EISDIR) via CLI, double-click or
    macOS open-file shows a localized error dialog (errOpenFailed in all 20 locales)
    instead of failing silently; session restore stays quiet (PR #154).

## [0.16.0] - 2026-09-23

### Added

- Slides: Ctrl+M inserts a slide, Tab cycles into and out of groups, JPEG export,
  a Slide Size dialog, groups that contain tables, whole-body notes formatting, and
  chartEx treemap/waterfall fallback rendering (PR #120).
- Shell: local crash diagnostics — the crash reporter writes minidumps to
  `userData/crash-dumps` (upload disabled, nothing leaves the machine) and the main
  process keeps a 1 MiB ring-buffer log at `userData/logs/main.log` (3 backups) with
  structured records for render-process-gone, child-process-gone and quit, so "the
  app closed itself" reports can be diagnosed on the user's machine (PR #138).

### Changed

- docs/COPILOT.md documents an honest per-format zip budget matrix (what each format
  fences on open and what remains unfenced), every row anchored to code (PR #119).
- Focus-ring tokens across pdf/docs/markdown raised to WCAG ≥3:1 contrast via alpha
  only (worst token 1.13 → 3.06; hue and geometry untouched), with a contrast
  invariant test (PR #122).
- Test infrastructure: local xvfb e2e pins byte-exact Carlito/Caladea 2013 CI font
  builds via `setup:e2e-fonts` (removes five false pixel-baseline failures) and
  sheets abort AI streams on tab teardown (PR #113); a tests-only coverage batch
  adds 20 tests over sort edge branches, SmartArt gallery UI, Home file-search
  wiring and mcp md/html suite isolation (PR #117).

### Fixed

- Sheets:
  - Legacy .xls files no longer turn Cyrillic strings into garbage — the BIFF8
    shared-string table is decoded as UTF-16 through a dedicated OLE2/BIFF reader
    instead of a single-byte codepage (user-reported; PR #126).
  - List validation pointing at a hidden sheet no longer blocks typing for ~20 s
    until the whole workbook streams in — the validator gates to VALID while the
    source range is not loaded and refreshes the sources after the sheet loads
    (user-reported; PR #127).
  - The Carlito-Regular/Bold.ttf 404 at every launch is gone — cell-font fallback
    uses static font imports instead of a bare runtime URL (PR #128).
  - Printing a wide sheet no longer surprises: a guard detects multi-strip jobs and
    paper that mismatches the UI locale (file Letter vs expected A4) and offers
    fit-to-width and a paper switch in the Print dialog; the file's print semantics
    are untouched (user-reported; PR #129).
  - .xls conversion fidelity: a second BIFF8 layout pass transfers merged cells,
    column widths and cell styles (FONT/XF/FORMAT/PALETTE into styles.xml), so forms
    stop falling apart; zip/BIFF5 output stays byte-identical (user-reported;
    PR #130).
  - A new workbook's grid grows on demand (Name Box/Go To, arrow keys at the edge;
    5,000×64 steps up to Excel's 1,048,576×16,384 cap) — typing below row 1000 no
    longer fails with "Range out of bounds"; new sheet tabs start at 1000×26
    (PR #132).
  - File → Print works on unmodified workbooks (it was gated on having pending
    edits, unlike the native menu/Ctrl+P path); save/save-as/export guards are
    untouched (PR #142).
  - Print tails: header/footer positions account for the Chromium inset,
    multi-section number formats round half-away like Excel, and General-format
    dates stay dates after recalc (PR #112).
- Docs:
  - A trailing CJK comma on a document-grid line is visible again — the
    punctuation-hang renderer plugin no longer zeroes the glyph advance and trusts
    Blink with a zero-advance draw; hanging is grouped by box overlap, the inked
    glyph is deterministic, and failed hangs are blacklisted (PR #140).
  - A corrupt (truncated XML) .docx is refused with a "file is damaged" error
    through the existing open-failure channel instead of silently opening as an
    empty document; all 32 valid corpus files still open without false refusals
    (PR #131).
  - Glyph metrics: U+25CB gets a measured fallback width, and numbered lists without
    w:ind no longer grow a phantom indent (PR #111).
  - F9 refreshes field caches under TRACK_IGNORE without writing revisions, and
    _Ref anchor ids are stamped from the document's free w:id pool (PR #114).
  - Footnote pagination on the dense-notes doc-06 document: the page-1 paragraph
    count now matches Word (PR #109).
- Slides:
  - Video export actually records: MediaRecorder receives a real
    `canvas.captureStream()` stream instead of a wrapper object — a live export
    produces a valid MP4 (ffprobe-verified under xvfb) where the feature silently
    failed since it landed (PR #133).
  - PDF export embeds the bundled fonts deterministically: the Calibri→Carlito
    bundle mapping no longer depends on system-installed fonts, and chart labels in
    the export SVG carry the same font stack the canvas draws — exported PDFs show
    only Carlito where LiberationSans used to appear (PR #137).
  - Recorder cancellation can no longer hang, per-slide PNG fit scales are restored,
    recovery escalates after an aborted quit, the live bridge restart is serialized
    per process under a token, Escape fires once, and desktop .pptx opens enforce
    the headless zip fences with int32-normalized declared sizes (PR #116).
- Shell:
  - Quit hardening: the session persists in before-quit, closed docs tabs park on
    about:blank, the dirty set is recomputed before the final close, quit-and-
    install goes through a full `app.quit()`, and titles truncate on code points
    without breaking surrogates (PR #118).
  - App-settings persistence is written by a single writer, closing the window where
    a burst of concurrent writes could drop keys in transit (fs.watch repro;
    PR #108).
- MCP:
  - `setHeadingLevel` to a level with no style in the document now survives save and
    reopen (direct `w:outlineLvl`, Word's native override) and reports a warning
    instead of silently keeping the old style (PR #134).
  - `apply_workbook_ops` validates A1 refs on write as well as read: "A0"-class and
    beyond-grid refs are rejected per operation with a clear error instead of
    poisoning the journal so every later save failed and unsaved edits were lost
    (PR #136).
  - Hardening batch: encrypted .docx is refused, insert markers must sit inside the
    body, rich-text booleans are validated, the zip-bomb budget counts entries, and
    `set_image_properties` validates inputs (PR #115).
  - Infra batch: zip64 u64 saturation is pinned by a test, a stat→open TOCTOU on
    xlsx opens is closed via rawBytes in the wire reply, negative declared zip sizes
    are normalized in the docx fence, and quit snapshots retry with a will-quit
    flush (PR #123).

### Security

- .ods conversion is fenced by the same central-directory zip budget as .xlsx
  before the converter spawns; legacy OLE2 .xls bypasses it as before (PR #110).
- JSZip declared sizes ≥2 GiB no longer wrap to negative int32 and slip past zip
  budgets — normalized at docx zip-load and the MCP size fence (PR #121).
- A hostile .xls can no longer crash the sidecar or drive it into an unbounded
  allocation: OLE2 headers are validated before calamine (geometry no valid
  [MS-CFB] container has), and the three calamine calls on the conversion path run
  under catch_unwind (PR #139).

### Performance

- Big documents (18k paragraphs): the per-transaction O(n²) line-factor DecorationSet
  rebuild is gone — opening is 10/10 runs ≤3 s (523–777 ms; the intermittent
  multi-minute stalls are gone), scrolling went 7.2 → 60 fps with zero long-frame
  gaps, and typing on giants dropped from 4.8 s to 1.1 s per keystroke (32 ms on
  normal documents; block virtualization is a follow-up) (PR #135).
- Closed tabs stop retaining renderer processes: a closed docs tab's webContents is
  now destroyed after the about:blank teardown navigation — over 60 real close
  cycles, orphan samples went 47/60 → 0/60 and the RSS envelope 1928 MB / 11
  processes → 615–637 MB / 4 (baseline 607 MB) (PR #141).
- CI wall time cut 10m56s → ~4m40s warm (2.34×): the pipeline is split into
  parallel jobs without needs chains (test ×6 with vitest shards, e2e ×2 Playwright
  shards) with electron/rust caches and build-once e2e artifacts (PR #124), then an
  affected-only matrix and a third e2e shard (PR #125).

<!-- Link refs to update at the bottom of CHANGELOG.md:
[Unreleased]: https://github.com/besliky/airy/compare/v0.16.0...HEAD
[0.16.0]: https://github.com/besliky/airy/compare/v0.15.0...v0.16.0
-->

## [0.15.0] - 2026-09-22

### Added

- Docs: TOC options support style sources (`\t`) with per-level page-number
  ranges; text-column controls; WordArt transform presets; Notes/Handouts PDF
  export layouts; alt text editing across element kinds.

### Fixed

- Rendering fidelity (product-level audit round): pages stopped double-counting
  the leader paragraph's top margin on natural breaks (content sat 13–16 px lower
  than Word) and adjacent paragraph margins now sum like Word instead of
  collapsing (content sat 8–10 px higher); CJK↔Latin autospacing carries Word's
  full 1/4 em gap; hyphenation degrades to manual when the runtime has no
  dictionaries.
- Slides: video export streams frames and file chunks (peak memory on a
  30-minute deck drops from ~6 GB to ~30 MB), recorder errors surface instead
  of truncating silently, and fly-in/out animations enter horizontally in
  PowerPoint; PDF export embeds the deck's fonts so text no longer falls back
  on machines without them.
- Sheets: hidden and filtered rows/columns no longer print (a filtered export
  used to leak the filtered-out data); conditional-formatting fills, font
  colors, and data bars print; manual page breaks are honored; collapsed
  borders no longer produce a blank trailing page; time formats truncate like
  Excel; text-to-columns clamps to the used range and confirms streamed
  destinations before overwriting.
- Shell: cancelling quit fully rolls back (window-close prompts latch, the
  live bridge restarts); the Home recents list updates live as tabs open.
- MCP: `setHeadingLevel` and `clearList` survive saves; `read_workbook` sees
  pending edits; combined live operations undo with exactly two `live_undo`
  calls; workbook opens enforce zip-bomb byte budgets on every format path.
- Engine: style definitions are patched surgically on Modify (numbering, tabs,
  and keep-with-next survive) with Word-style explicit offsets; built-in
  styles no longer get customStyle markers.

### Changed

- Test infrastructure: editors mounted in docs tests are tracked and destroyed
  behind a lifecycle contract (fixes an intermittent CI failure); local e2e
  runs headless under xvfb via `npm run test:e2e:xvfb`.

## [0.14.0] - 2026-09-21

### Added

- Docs: TOC options gain style-source mapping — fields built from named styles
  (`\t`) alongside outline levels, with per-level page-number ranges.
- Slides: Escape layering, popover close, and split-caret keyboard access; effect
  option disabled-reasons; per-transition duration clamps with localized hints.

### Fixed

- Sheets: Text-to-Columns clamps to the used range and refuses >50k-row selections
  with a clear message; print tile continuations inherit anchor cell styles so
  gridlines and fills survive page breaks; page order round-trips through the
  file and preselects in the print dialog; the visual-undo registry is cleared on
  teardown.
- Slides: video export records encoder errors instead of truncating silently and
  restores the active-export flag on renderer crashes; the render-phase cancel
  works; per-slide aspect fit for mixed-size decks; fly-in/out animations no
  longer add a diagonal offset in PowerPoint; split animation tokens serialize in
  the OOXML-legal form.
- Shell: cancelling quit now fully rolls back (sheets shutdown flag, live bridge
  restart) and restores per-window sessions; window-close prompts latch; restore
  windows cascade instead of stacking; video-export focus stays trapped in the
  dialog during recording.
- MCP: geometry NaN/negative guards; temp-file orphans cleaned on double failure;
  zip-bomb size fences on open (stat-first 512 MiB and pptx central-directory
  budgets).
- Docs: note spill budgets survive multiple footnotes and carry remainders;
  built-in style inventory completed (no more customStyle markers on Word
  built-ins); explicit offsets write complex-script twins; workbook sync keeps
  cell-less rows and non-A1 dimensions.
- Tests: editors mounted in docs tests are tracked and destroyed behind a
  lifecycle contract (fixes an intermittent unhandled-error CI failure); crypto
  and retry-pacing seams bring the serial suite back under 140s.

### Changed

- Shared tracked-editor test helper replaces per-file cleanup idioms (14 files
  migrated).

## [0.13.0] - 2026-09-19

### Added

- Shell: Move tab to a new window — tabs carry their live editor session between
  windows, multi-window session restore, and window-scoped tab routing for background
  opens. macOS gains a manual "Check for Updates" dialog linking to the releases page.
- MCP: headless slides deck tools (open/read/insert/save for .pptx) and read-only PDF
  text extraction — every Airy format now has headless tools.
- Slides: alt text editing for all element kinds with a collision-free chart marker;
  transition directions and exact durations plus animation directions (Effect
  Options); text columns, WordArt transform presets, and Notes/Handouts PDF export
  layouts.
- Sheets: print scopes (selection / active sheets / entire workbook with per-sheet
  headers, page order, and collation), a full Text-to-Columns wizard (multi-delimiter,
  fixed-width, column types, destination), and the outline +/- gutter with undo and
  Outline Settings.
- Docs: pagination compatibility profiles — Word parity 97.3% page-starts (page counts
  25/25) and a LibreOffice profile at 95.7% with hard test gates on both.

### Fixed

- Styles: Modify Style patches definitions surgically (keepNext/numbering/tabs
  survive) and supports explicit offsets (w:val="0") against inherited facets.
- Docs: TOC updates preserve section breaks; picture compression keeps the crop rect
  unless asked; shadow effects merge into existing effect lists; chart workbook sync
  preserves authored formulas; all TOC/TOF fields update together; wildcard search and
  compare hardening from the audit cycle.
- Slides: PDF export scopes SVG ids per slide (no more cross-slide clip/gradient
  bleed), flipped groups render text correctly, and workbook-wide printing labels
  pages with the owning sheet name.
- Shell: cancelling quit restores per-window sessions; tab detach/adopt no longer
  leaks listeners; window-scoped path dedup allows one file in two windows.
- MCP: case-insensitive replace is index-safe around U+0130; appends no longer
  fabricate phantom lines; renamed workspace roots fail loudly.
- Bridge: closed connections stop dispatching; dribbled chunks coalesce; oversized
  answers return a typed error; the drain deadline cannot be reset by peer data.

### Changed

- Modal dialog semantics (focus trap, Escape, role) now enforced by per-app coverage
  contracts across docs, sheets, and slides; the outline gutter is keyboard-accessible.
- Test wall time back under 135s despite +600 tests since v0.12.0 (scaled fixture
  budgets, collapsed retry pacing); docs renderer entry −43% via lazy locale
  dictionaries carries into CI bundle gates.

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

[Unreleased]: https://github.com/besliky/airy/compare/v0.15.0...HEAD
[0.15.0]: https://github.com/besliky/airy/compare/v0.14.0...v0.15.0
[0.14.0]: https://github.com/besliky/airy/compare/v0.13.0...v0.14.0
[0.13.0]: https://github.com/besliky/airy/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/besliky/airy/compare/v0.11.0...v0.12.0
[0.11.0]: https://github.com/besliky/airy/compare/v0.10.0...v0.11.0
[0.10.0]: https://github.com/besliky/airy/compare/v0.9.3...v0.10.0
[0.9.3]: https://github.com/besliky/airy/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/besliky/airy/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/besliky/airy/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/besliky/airy/releases/tag/v0.9.0
