# Changelog

All notable changes to `office-cli` are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project adheres
to [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-05-06

Initial public release — 115 atomic commands across four Office formats.

### Added

**Universal commands (auto-detect format):**
- `info` — one-line format + size + primary counter probe.
- `meta` — full metadata (title/author/pages/slides/sheets) for any format.
- `extract-text` — plain text dump from docx/xlsx/pptx/pdf.

**Excel (.xlsx) — 45 commands:**
- `excel sheets` — list worksheets with row/col counts.
- `excel read` — read sheet contents (full sheet, range, limit, with-headers, typed).
- `excel cell` — read a single cell value (lightweight).
- `excel write` — write one or many cells.
- `excel append` — append rows to a sheet.
- `excel search` — find cells containing a keyword (case-insensitive).
- `excel create` — create a new workbook from a JSON spec.
- `excel to-csv` / `excel from-csv` — CSV import/export (with UTF-8 BOM support).
- `excel to-json` / `excel from-json` — JSON import/export.
- `excel info` — quick workbook overview with optional preview rows.
- `excel rename-sheet` / `excel delete-sheet` / `excel copy-sheet` — sheet lifecycle.
- `excel hide-sheet` / `excel show-sheet` — sheet visibility.
- `excel style` — apply font/fill/border/alignment/number-format to a cell range.
- `excel batch-style` — apply multiple styles in a single open/save cycle.
- `excel cell-style` — read style properties of a cell.
- `excel default-font` / `excel set-default-font` — workbook default font.
- `excel insert-rows` / `excel insert-cols` — insert blank rows or columns.
- `excel delete-rows` / `excel delete-cols` — delete rows or columns.
- `excel sort` — sort a range by a column (ascending/descending).
- `excel freeze` — freeze panes at a cell reference.
- `excel merge` / `excel unmerge` — merge or unmerge cell ranges.
- `excel set-col-width` / `excel set-row-height` — set column width or row height.
- `excel chart` / `excel multi-chart` — single and multi-series charts.
- `excel add-image` — insert an image anchored at a cell.
- `excel cond-format` — conditional formatting (cell, top, bottom, average, duplicate, unique, formula).
- `excel data-bar` / `excel icon-set` / `excel color-scale` — enhanced conditional formats.
- `excel auto-filter` — set auto-filter on a header row.
- `excel formula` / `excel column-formula` — formula generation (SUM, AVERAGE, VLOOKUP, IF, SUMIF, CONCAT, CUSTOM).
- `excel copy` / `excel copy-range` / `excel fill-range` — workbook and range operations.
- `excel validation` — data validation rules (dropdown, numeric, date, custom).
- `excel hyperlink` / `excel get-hyperlink` — cell hyperlinks.

**Word (.docx) — 25 commands:**
- `word read` — read paragraphs (paragraphs | markdown | text, with tables support).
- `word replace` — find/replace text (single or batch via --pairs).
- `word search` — keyword search across paragraphs and table cells.
- `word meta` — document metadata.
- `word stats` — paragraph/heading/word/character/line counts.
- `word headings` — heading outline (Title + Heading1..6).
- `word images` — extract embedded images.
- `word create` — create a brand-new .docx document.
- `word add-paragraph` / `word add-heading` — append paragraphs and headings.
- `word add-table` — append a table from JSON rows.
- `word add-image` — insert an inline image.
- `word add-page-break` — insert a page break.
- `word delete` — remove a body element by index.
- `word insert-before` / `word insert-after` — insert content at a position.
- `word update-table-cell` — modify a single table cell.
- `word update-table` — batch update multiple table cells.
- `word add-table-rows` / `word delete-table-rows` — table row operations.
- `word add-table-cols` / `word delete-table-cols` — table column operations.
- `word merge` — combine multiple .docx files into one.
- `word style` — apply paragraph/run formatting (bold, italic, font, color, alignment, spacing).
- `word style-table` — apply cell/run formatting to table ranges.

**PowerPoint (.pptx) — 17 commands:**
- `ppt read` — read slide outlines (titles + bullets, markdown, text, with notes).
- `ppt replace` — find/replace text across every slide.
- `ppt meta` — presentation metadata.
- `ppt count` — slide count.
- `ppt outline` — lightweight title-only outline.
- `ppt images` — extract embedded images.
- `ppt create` — create a brand-new .pptx presentation.
- `ppt add-slide` — append a slide with title and bullets.
- `ppt set-content` — overwrite a slide's text content.
- `ppt set-notes` — set/replace speaker notes for a slide.
- `ppt delete-slide` — remove a slide by number.
- `ppt reorder` — reorder slides (e.g. "3,1,2").
- `ppt add-image` — insert an image into a specific slide.
- `ppt build` — create a complete deck from a JSON spec (with template and image support).
- `ppt layout` — read the shape tree (position, size, type, placeholder, text, font info).
- `ppt set-style` — modify text styling within a specific shape.
- `ppt add-shape` — insert new shapes (text-box, rect, ellipse, line, arrow).

**PDF — 22 commands:**
- `pdf read` — extract text per-page or as a single string.
- `pdf search` — find text occurrences with context snippets.
- `pdf pages` — page count with optional dimensions.
- `pdf info` — full metadata (encryption, signatures, dimensions, ...).
- `pdf create` — create a new PDF from a JSON spec.
- `pdf replace` — find and replace text in content streams.
- `pdf add-text` — overlay text at specific coordinates.
- `pdf merge` — concatenate multiple PDFs.
- `pdf split` — split into chunks of N pages.
- `pdf trim` — keep only listed pages.
- `pdf watermark` — add text watermark.
- `pdf stamp-image` — stamp an image watermark on pages.
- `pdf rotate` — rotate pages by 90/180/270 degrees.
- `pdf optimize` — reduce file size.
- `pdf encrypt` / `pdf decrypt` — password-protect / strip encryption.
- `pdf extract-images` — extract embedded images.
- `pdf bookmarks` — read PDF outline/bookmarks tree.
- `pdf add-bookmarks` — add bookmarks to a PDF.
- `pdf reorder` — reorder pages (e.g. "3,1,2,4").
- `pdf insert-blank` — insert blank pages at a position.
- `pdf set-meta` — set PDF metadata (title, author, subject, keywords).

**Operational:**
- `doctor` — runtime environment, config validation, and optional tool check.
- `reference` — all commands/flags as structured Markdown.
- `setup` — create or verify configuration file.
- `install-skill` — install bundled AI Agent Skill.

**Infrastructure:**
- Cobra command tree with global `--json` / `--quiet` / `--dry-run` / `--force` flags.
- Three-tier permission system (read-only / write / full) via `~/.office-cli/config.json`.
- JSONL audit logger under `~/.office-cli/audit/` with monthly rotation and sensitive arg stripping.
- Structured error codes (`errorCode` + `hint` + `file`) for AI Agent consumption.
- AI Agent Skill bundled at `skills/office-cli/SKILL.md` with format-specific docs.
- npm distribution via `@fatecannotbealtered-/office-cli` with postinstall binary download and SHA-256 verification.
- CI (3 OS matrix × go vet / gofmt / test / E2E) + Release (goreleaser) + npm publish workflows.
- Unit tests for all engine packages (excel, word, ppt, pdf) and E2E tests for all 115 commands.
