---
name: office-cli
description: Local Office documents (Word, Excel, PPT, PDF) toolkit. 115 atomic commands: read, search, edit, create, style, sort, merge, split, watermark, chart, and convert .docx/.xlsx/.pptx/.pdf without any cloud round-trip. Always use --json when parsing programmatically.
metadata: {"agent":{"emoji":"📄","requires":{"bins":["office-cli"]}}}
---

# office-cli

Offline CLI for the four most most common Office document formats: **Word (.docx)**, **Excel (.xlsx)**, **PowerPoint (.pptx)** and **PDF**. Designed first-class for AI Agents.

> Install CLI: `npm install -g @fatecannotbealtered-/office-cli`
>
> Install Skill: `npx skills add fatecannotbealtered/office-cli -y -g`

## When to use this skill

Use whenever the user wants to:

- **Read** a local Office document (extract text, find a value, summarise structure)
- **Edit** a local Office document (change text, add rows, replace strings, add a watermark)
- **Create** a new document from scratch (spreadsheet, report, slide deck)
- **Convert** between formats (xlsx → csv, docx → markdown)
- **Restructure** a PDF (merge, split, trim, reorder, add bookmarks)
- **Inspect metadata** (author, page count, slide count, sheet list)
- **Style / format** Excel cells (fonts, colours, borders, charts, conditional formatting)

Do NOT use this skill for:
- Cloud documents (Google Docs, Office 365 online, Lark Docs) — use the relevant cloud skill instead
- Image-only PDFs requiring OCR — `pdf read` returns content-stream text only

## Prerequisites

There is **nothing to authenticate**. office-cli works fully offline using built-in Go libraries.

Run `office-cli doctor --json` once to confirm the binary runs and to see whether optional external tools (`pandoc`, `libreoffice`) are installed; they are only needed for advanced format conversions and are NOT required for the commands documented below.

## Setup & Permissions

```bash
# Create default config (idempotent — safe to call repeatedly)
office-cli setup

# Verify configuration
office-cli doctor --json
```

Config file: `~/.office-cli/config.json` (created by `setup`).

### Permission system

office-cli enforces a three-tier permission model. Default is the most restrictive level.

| Level | Description |
|-------|-------------|
| `read-only` (default) | Read operations only |
| `write` | + modify operations (write cells, append rows, replace text, style, sort, create documents, add/delete/insert slides/paragraphs/tables, update table cells, merge docs, merge/split/reorder PDFs) |
| `full` | + destructive operations (delete sheets, encrypt/decrypt PDFs) |

**To change permissions:** edit `~/.office-cli/config.json` directly:
```json
{ "permissions": { "mode": "write" } }
```

**CI/CD override:** set `OFFICE_CLI_PERMISSIONS=write` (or `full`).

Commands that require higher permissions will exit with code 5 (`PERMISSION_DENIED`) if the configured mode is insufficient.

## Conventions

- **First positional argument is always the file path.** Quote paths containing spaces.
- **Add `--json` when parsing output programmatically.** Without it, output is human-formatted (colored tables, ANSI codes) and unsuitable for automation.
- **Add `--quiet` to suppress non-result stdout.** Combine with `--json` for clean pipe-friendly output.
- **Add `--dry-run` before write operations** to preview the action without touching the filesystem.
- **`--force` skips confirmation prompts.** Use only when the user explicitly authorises destructive steps.
- All write operations are recorded in `~/.office-cli/audit/` (JSONL, one file per month). Set `OFFICE_CLI_NO_AUDIT=1` to disable.

## Universal commands (auto-detect format)

When the file format is not known up front (e.g. user-uploaded `report.???`),
prefer these dispatchers — they sniff the extension and route to the right engine:

```bash
office-cli info  any.docx --json    # one-line probe (format + size + primary counter)
office-cli meta  any.pdf  --json    # full FlatDocMeta: title/author/pages/slides/sheets
office-cli extract-text any.pptx --json   # plain text dump from any of docx/xlsx/pptx/pdf
```

`info` is the cheapest call when you only need to know "what is this file"; use
it before spawning a series of format-specific reads.

## Format-specific commands

Each format has a dedicated reference file with all commands and examples:

| Format | Commands | Reference file | ~Tokens | Key capabilities |
|--------|----------|----------------|---------|------------------|
| **Excel (.xlsx)** | 45 | [`docs/excel.md`](docs/excel.md) | ~4k | read/write cells, style, sort, freeze, merge, charts, cond-format, auto-filter, images, CSV, sheet lifecycle, column formulas (VLOOKUP/IF/SUMIF/CONCAT), JSON conversion, range ops, multi-series charts, data bars, icon sets, color scales, hyperlinks, default font, cell style read |
| **Word (.docx)** | 25 | [`docs/word.md`](docs/word.md) | ~3k | create, add paragraphs/headings/tables/images/page breaks, read (with tables), replace, search, delete, insert-before/after, update-table-cell, update-table (batch), add/delete table rows/cols, merge, metadata, headings, images, style (paragraph/run formatting), style-table (cell formatting + batch) |
| **PowerPoint (.pptx)** | 17 | [`docs/ppt.md`](docs/ppt.md) | ~3k | create, add slides, set content/notes, delete/reorder slides, add images, read, replace, metadata, build from JSON spec (with images + template support), layout (shape tree), set-style, add-shape |
| **PDF** | 22 | [`docs/pdf.md`](docs/pdf.md) | ~3k | read, search, create, replace, merge, split, trim, watermark, stamp, rotate, optimize, encrypt/decrypt, bookmarks, reorder, insert blank, set meta, add-text |

**Loading strategy:** Read the format-specific file only when the user's request targets that format. Start with this file for overview and routing. Each file is ~3–4k tokens — load only what you need.

## Workflow examples

### Fill a Word template from a JSON record

```bash
# 1. Inspect the template to confirm placeholders are intact (no run splits)
office-cli word read template.docx --format text --quiet | head -50

# 2. Apply replacements
office-cli word replace template.docx --pairs '[
  {"find":"{{customer_name}}","replace":"Acme Inc."},
  {"find":"{{contract_id}}","replace":"C-2024-001"},
  {"find":"{{total}}","replace":"$98,765.00"}
]' --output contract.docx
```

### Edit a Word document's structure (tables, headings, paragraphs)

```bash
# 1. Discover the document structure (paragraphs + tables in order)
office-cli word read report.docx --with-tables --json

# 2. Delete an unwanted element (0-based index from step 1)
office-cli word delete report.docx --index 3 --output report-edited.docx

# 3. Update a table cell (table-index is the body-element index)
office-cli word update-table-cell report-edited.docx --table-index 1 --row 2 --col 1 --value "PASS" --output report-final.docx

# 4. Insert a new section heading after element 0
office-cli word insert-after report-final.docx --index 0 --type heading --text "Executive Summary" --level 1 --output report-final.docx

# 5. Merge two documents
office-cli word merge cover.docx body.docx --output full-report.docx
```

### Roll up several spreadsheets into one summary

```bash
# 1. Discover what's in each
for f in q1.xlsx q2.xlsx q3.xlsx q4.xlsx; do
  office-cli excel sheets "$f" --json
done

# 2. Read totals (assume each has a 'Summary' sheet with a 'Total' cell at B10)
office-cli excel read q1.xlsx --sheet Summary --range B10:B10 --json --quiet
# ...repeat for q2..q4...

# 3. Build the summary workbook
office-cli excel create year-summary.xlsx --spec '{
  "sheets":[{"name":"Yearly","rows":[
    ["Quarter","Total"],
    ["Q1",111111],
    ["Q2",222222],
    ["Q3",333333],
    ["Q4",444444]
  ]}]
}'
```

### Build an executive PDF from many sources

```bash
# 1. Trim cover pages from the input reports
office-cli pdf trim a.pdf --pages "2-" --output a.body.pdf
office-cli pdf trim b.pdf --pages "2-" --output b.body.pdf

# 2. Merge in the desired order
office-cli pdf merge --input cover.pdf --input a.body.pdf --input b.body.pdf --output exec-summary.pdf

# 3. Add a watermark
office-cli pdf watermark exec-summary.pdf --text "INTERNAL" --output exec-summary.wm.pdf
```

### Search a PDF and create a summary

```bash
# 1. Check if a contract mentions a specific clause
office-cli pdf search contract.pdf "termination" --json

# 2. Create a one-page summary PDF from findings
office-cli pdf create summary.pdf --spec '{"pages":[{"text":"Contract Review\n\nKey findings:\n- Termination clause found on page 5\n- No penalty clause detected"}],"title":"Contract Review"}' --json
```

### Replace placeholders in a PDF template

```bash
# 1. Create a PDF template
office-cli pdf create template.pdf --spec '{"pages":[{"text":"Invoice #{{INV_NUM}}\n\nClient: {{CLIENT_NAME}}\nAmount: {{AMOUNT}}"}]}' --json

# 2. Fill in the template
office-cli pdf replace template.pdf --pairs '[
  {"find":"{{INV_NUM}}","replace":"INV-2026-001"},
  {"find":"{{CLIENT_NAME}}","replace":"Acme Corp"},
  {"find":"{{AMOUNT}}","replace":"$12,345.00"}
]' --output invoice-001.pdf --json
```

### Stamp text overlays on a PDF

```bash
# 1. Create a report
office-cli pdf create report.pdf --spec '{"pages":[{"text":"Quarterly Results\n\nRevenue: $1.2M\nProfit: $340K"},{"text":"Appendix\n\nDetailed tables..."}],"title":"Q1 Report"}' --json

# 2. Add "DRAFT" overlay in red on every page
office-cli pdf add-text report.pdf --text "DRAFT" --x 200 --y 400 --font-size 48 --color "1 0 0" --output report.draft.pdf --json

# 3. Add page-specific overlays via --spec
office-cli pdf add-text report.pdf --spec '[
  {"text":"CONFIDENTIAL","x":150,"y":700,"fontSize":24,"color":"1 0 0"},
  {"text":"Internal Use Only","x":72,"y":50,"fontSize":10,"pages":"1"}
]' --output report.stamped.pdf --json
```

### Create a PDF with typography features

```bash
# 1. Create a PDF with serif font and centered text
office-cli pdf create report.pdf --spec '{
  "font": "Times",
  "fontSize": 14,
  "align": "center",
  "pages": [
    {"text": "Annual Report", "bold": true, "fontSize": 24},
    {"text": "Executive Summary\n\nThis report covers..."},
    {"text": "Financial Highlights", "bold": true, "underline": true}
  ]
}' --json

# 2. Create a PDF with mixed fonts and styles
office-cli pdf create mixed.pdf --spec '{
  "pages": [
    {"text": "Sans-serif Title", "font": "Helvetica", "bold": true},
    {"text": "Serif body text", "font": "Times"},
    {"text": "Monospace code", "font": "Courier", "align": "left"},
    {"text": "Right-aligned note", "align": "right", "italic": true}
  ]
}' --json

# 3. Add underlined text overlay with serif font
office-cli pdf add-text report.pdf --text "Approved" --font Times --bold --underline --x 200 --y 400 --output report.approved.pdf --json
```

### Build a PowerPoint deck from a JSON spec

```bash
# 1. Create a deck in one shot (reduces N round-trips to 1)
office-cli ppt build deck.pptx --spec '{
  "title": "Q2 Review",
  "author": "Alice",
  "slides": [
    {"title": "Overview", "bullets": ["Revenue up 20%", "3 new products launched"]},
    {"title": "Roadmap", "bullets": ["Q3: scale infrastructure", "Q4: expand to APAC"], "notes": "Mention the hiring plan"},
    {"title": "Q&A"}
  ]
}' --json

# 2. Build from a corporate template (preserves slide master, theme, layouts)
office-cli ppt build deck.pptx --template company.pptx --spec '{
  "title": "Q2 Report",
  "slides": [
    {"title": "Overview", "bullets": ["Revenue up 20%"]},
    {"title": "Details"},
    {"title": "Q&A"}
  ]
}' --json

# 3. Build with images embedded in specific slides
office-cli ppt build deck.pptx --spec '{
  "slides": [
    {"title": "Chart", "image": "chart.png", "width": 6000000, "height": 4000000},
    {"title": "Summary", "bullets": ["Key takeaway"]}
  ]
}' --json

# 4. For large specs, use @file to avoid shell escaping issues:
cat > /tmp/deck-spec.json << 'EOF'
{
  "title": "Annual Review",
  "author": "Alice",
  "slides": [
    {"title": "Revenue", "bullets": ["$1.2M ARR", "40% YoY growth"]},
    {"title": "Product", "bullets": ["3 new features", "99.9% uptime"]},
    {"title": "Team", "bullets": ["12 engineers", "2 new hires"]},
    {"title": "Roadmap", "bullets": ["Q3: scale", "Q4: expand"], "notes": "Mention hiring plan"},
    {"title": "Q&A"}
  ]
}
EOF
office-cli ppt build deck.pptx --spec @/tmp/deck-spec.json --json

# 5. Multi-step: replace placeholders after build
office-cli ppt build deck.pptx --spec '{"title":"__TITLE__","slides":[{"title":"Overview","bullets":["__METRIC__"]}]}'
office-cli ppt replace deck.pptx --find "__TITLE__" --replace "Q2 Review" --output deck.final.pptx
```

### Layout-aware PPT editing (inspect → style → create shapes)

```bash
# 1. Inspect the shape tree to understand visual layout
office-cli ppt layout deck.pptx --slide 1 --json
# Returns: shape index, type, name, placeholder, position/size (EMU), text, font info

# 2. Style an existing shape (change font, color, alignment)
office-cli ppt set-style deck.pptx --slide 1 --shape 0 --font-size 4800 --bold --color "FF0000" --output deck.pptx
office-cli ppt set-style deck.pptx --slide 1 --shape 1 --align center --output deck.pptx

# 3. Add new shapes to a slide
office-cli ppt add-shape deck.pptx --slide 1 --type text-box --x 500000 --y 200000 --width 4000000 --height 1000000 --text "New Section" --font-size 2400 --bold --fill "E8E8E8" --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type rect --x 0 --y 7000000 --width 9144000 --height 500000 --fill "4472C4" --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type arrow --x 1000000 --y 5000000 --width 3000000 --height 0 --output deck.pptx

# 4. Verify changes
office-cli ppt layout deck.pptx --slide 1 --json
```

## Guardrails

- Always preview write operations with `--dry-run` first when the user has not seen the source file
- Use `--force` only when the user explicitly authorises overwriting / skipping confirmation
- For replacements that do not produce hits, do NOT loop endlessly — switch to placeholder-token strategy or report back to the user
- `pdf trim` and `pdf split` write new files; the input PDF is never mutated
- `word replace`, `ppt replace`, and `pdf replace` write a separate output file by default (`<input>.replaced.<ext>`); use `--output` to override
- All write commands are audited in `~/.office-cli/audit/`; honour `OFFICE_CLI_NO_AUDIT=1` if the user wants privacy
- `excel create` refuses to overwrite an existing file unless `--force` is set

### Replace limitations (important)

`word replace`, `ppt replace`, and `pdf replace` only match text that lives **entirely within a single XML text element**:

| Format | Limitation |
|--------|-----------|
| Word | Text must be within a single `<w:t>` element. Words split across runs (font boundaries, spell-check) are NOT matched. |
| PPT | Text must be within a single `<a:t>` element. |
| PDF | Only parenthesized string literals in Tj operators are matched. Hex strings `<...>` and text split across operators are NOT matched. |

When `hits` is 0, the JSON response includes a `warning` field explaining this. Do NOT retry with the same pairs — instead, report the limitation to the user.

### Using @file for large JSON specs

For `--spec`, `--pairs`, and `--spec` flags that accept JSON, prefer `@file` over inline JSON when the payload exceeds ~500 characters. This avoids shell escaping issues:

```bash
# Write spec to a temp file, then reference it
office-cli ppt build deck.pptx --spec @/tmp/spec.json --json
office-cli word replace doc.docx --pairs @/tmp/pairs.json --output out.docx --json
office-cli pdf add-text report.pdf --spec @/tmp/overlays.json --output out.pdf --json
```

## Global flags

- `--json` — Machine-readable output (results on stdout, errors on stderr)
- `--quiet` — Suppress non-JSON stdout (combine with `--json` for pipe-friendly output)
- `--dry-run` — Preview write commands without touching the filesystem
- `--force` — Skip confirmation prompts and allow overwrites

## JSON output schemas

### Excel sheet (from `excel sheets`)

```json
{
  "name": "Q1",
  "index": 0,
  "rows": 42,
  "cols": 8,
  "dimension": "A1:H42"
}
```

### PDF page (from `pdf read`)

```json
{
  "page": 5,
  "text": "Quarterly results...",
  "wordCount": 312
}
```

### Word paragraph (from `word read --format paragraphs`)

```json
{
  "index": 17,
  "style": "Heading2",
  "text": "Background"
}
```

### PPT slide (from `ppt read`)

```json
{
  "index": 3,
  "title": "Roadmap",
  "bullets": ["Q1: foundation", "Q2: GA"],
  "notes": "Skip if running short",
  "text": "Roadmap\nQ1: foundation\nQ2: GA"
}
```

### PPT shape (from `ppt layout`)

```json
{
  "index": 0,
  "type": "sp",
  "name": "Title 1",
  "ph": "title",
  "x": 0, "y": 0, "w": 9144000, "h": 1371600,
  "text": "Q2 Review",
  "paragraphs": [
    {
      "align": "l",
      "runs": [{"text": "Q2 Review", "fontSize": 3600, "bold": true, "color": "003366"}]
    }
  ]
}
```

### Error response

```json
{
  "error": "sheet not found: Marketing",
  "errorCode": "NOT_FOUND",
  "hint": "The requested sheet/page/slide does not exist in this document",
  "file": "sales.xlsx"
}
```

### Error codes

| Code | Hint |
|------|------|
| `FILE_NOT_FOUND` | Verify the file path exists and is reachable from the current working directory |
| `INVALID_FORMAT` | The file extension or content does not match the expected format |
| `CORRUPTED_FILE` | The document is corrupted or password-protected |
| `PERMISSION_DENIED` | OS-level permission issue (read/write denied or file in use) |
| `VALIDATION_ERROR` | Bad command arguments or invalid spec content |
| `NOT_FOUND` | Sheet/page/slide does not exist in this document |
| `ENGINE_ERROR` | Internal error from a document engine; re-run with `--json` for details |
| `TOOL_MISSING` | Optional external tool (pandoc/libreoffice) is missing |
| `UNKNOWN_ERROR` | Catch-all bucket; re-run with `--json` and inspect details |

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 2 | Bad arguments |
| 4 | File / sheet / page / slide not found |
| 5 | Permission denied |
| 7 | Engine / processing error |

## Audit logging

Every write command (write, append, create, replace, merge, split, trim, watermark, to-csv, style, sort, freeze, add-*, set-*, delete-*, reorder, insert-blank, stamp-image, set-meta) is logged to `~/.office-cli/audit/audit-YYYY-MM.jsonl`, one JSON object per line:

```json
{"ts":"2026-05-04T04:30:12Z","cmd":"office-cli excel write","args":["sales.xlsx","--ref","B5","--value","1234.56"],"exit":0,"ms":42}
```

| Env var | Default | Description |
|---------|---------|-------------|
| `OFFICE_CLI_HOME` | `~/.office-cli` | Override the home directory |
| `OFFICE_CLI_NO_AUDIT` | (unset) | Set `1` to disable audit logging |
| `OFFICE_CLI_AUDIT_RETENTION_MONTHS` | `3` | Auto-delete audit files older than N months (`0` = keep forever) |

## Self-description

```bash
office-cli reference   # All commands, subcommands and flags as Markdown
office-cli doctor      # Environment + optional tool check
```
