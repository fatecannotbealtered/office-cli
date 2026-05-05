# office-cli

Local Office documents (Word, Excel, PPT, PDF) toolkit for humans and AI Agents.

`office-cli` lets you **read, edit, search, merge, split, watermark and convert** Office documents from a single, JSON-friendly command line. There is **no cloud, no auth, no telemetry** — every operation runs on your machine using built-in Go libraries.

The CLI is intentionally designed to pair with a Skill (`skills/office-cli/SKILL.md`) so AI Agents can drive Office workflows safely and predictably.

English | [简体中文](./README_zh.md)

## Highlights

| Format | Commands | Read | Write | Key capabilities |
|--------|----------|------|-------|------------------|
| **Excel (.xlsx)** | 45 | sheets, read, cell, search, info, cell-style, default-font | write, append, create, to-csv, from-csv, to-json, from-json, style, sort, freeze, merge/unmerge, chart, multi-chart, data-bar, icon-set, color-scale, cond-format, auto-filter, add-image, insert/delete rows & cols, col-width, row-height, hide/show sheet, rename/copy/delete sheet, formula, column-formula, copy, copy-range, fill-range, validation, hyperlink | Full spreadsheet control |
| **Word (.docx)** | 25 | read, meta, stats, headings, images, search | create, replace, add-paragraph, add-heading, add-table, add-image, add-page-break, delete, insert-before/after, update-table-cell, update-table, add/delete table rows/cols, merge, style, style-table | Create & edit documents |
| **PowerPoint (.pptx)** | 17 | read, count, outline, meta, images, layout | create, replace, add-slide, set-content, set-notes, delete-slide, reorder, add-image, build, set-style, add-shape | Create & manage decks |
| **PDF** | 22 | read, pages, info, bookmarks, extract-images, search | merge, split, trim, watermark, stamp-image, rotate, optimize, reorder, insert-blank, add-bookmarks, set-meta, encrypt, decrypt, create, replace, add-text | Full PDF lifecycle |
| **Universal** | 3 | info, meta, extract-text | — | Auto-detect format |

**Total: 115 atomic commands** — every command does exactly one thing. AI Agents compose freely.

### AI-friendly by design

| Feature | Description |
|---------|-------------|
| `--json` | Machine-readable flat JSON output (stdout), errors to stderr |
| `--quiet` | Suppress non-JSON stdout — pipe-friendly |
| `--dry-run` | Preview write operations without touching the filesystem |
| `--force` | Skip interactive confirmations for CI/Agent automation |
| `errorCode` + `hint` | Stable structured error responses — never guess |
| Flat schemas | Token-efficient: `FlatSheet`, `FlatPage`, `FlatSlide`, `FlatParagraph` |
| `office-cli reference` | Every command + flag as Markdown — Agent self-discovery |
| Audit log | Every write command logged to `~/.office-cli/audit/audit-YYYY-MM.jsonl` |

## Install

### Via npm (recommended)

```bash
# Install the CLI binary
npm install -g @fatecannotbealtered-/office-cli

# Install the AI Agent Skill (for Cursor / Claude Code / OpenClaw / etc.)
npx skills add fatecannotbealtered/office-cli -y -g
```

The post-install script downloads the right prebuilt binary for your OS/arch and verifies the SHA-256 checksum.

### From source

```bash
git clone https://github.com/fatecannotbealtered/office-cli.git
cd office-cli
make build
./bin/office-cli doctor
```

## Quick tour

```bash
# First-time setup (creates ~/.office-cli/config.json)
office-cli setup

# Universal — sniff format and dispatch
office-cli info  any.docx --json          # one-line probe
office-cli meta  any.pdf  --json          # full metadata
office-cli extract-text any.pptx --json   # plain text from any format

# Excel — 30 commands for complete spreadsheet control
office-cli excel sheets sales.xlsx --json
office-cli excel read sales.xlsx --range A1:D10 --json
office-cli excel cell sales.xlsx B2 --sheet Sales --json
office-cli excel append sales.xlsx --rows '[["Alice", 100], ["Bob", 200]]'
office-cli excel from-csv data.csv --output data.xlsx
office-cli excel to-csv sales.xlsx --sheet Sales --output sales.csv --bom
office-cli excel rename-sheet workbook.xlsx --from Old --to New
office-cli excel copy-sheet workbook.xlsx --from Template --to Backup
office-cli excel style sales.xlsx --sheet Sales --range A1:D1 --bold --bg-color "4472C4" --text-color "FFFFFF"
office-cli excel sort sales.xlsx --sheet Sales --range A1:D100 --by-col 2 --ascending
office-cli excel freeze sales.xlsx --sheet Sales --cell A2
office-cli excel chart sales.xlsx --sheet Sales --cell E2 --type bar --data-range "A1:B10" --title "Revenue"
office-cli excel merge sales.xlsx --sheet Sales --range A1:D1

# Word — create, write, and read
office-cli word create report.docx --title "Q2 Report" --author "Alice"
office-cli word add-heading report.docx --text "Introduction" --level 1 --output report.docx
office-cli word add-paragraph report.docx --text "Body text..." --output report.docx
office-cli word add-table report.docx --rows '[["Name","Score"],["Alice",95]]' --output report.docx
office-cli word add-image report.docx --image chart.png --width 400 --height 300 --output report.docx
office-cli word read report.docx --format markdown
office-cli word replace template.docx --find "{{name}}" --replace "Alice"
office-cli word headings report.docx --json
office-cli word stats report.docx --json

# PowerPoint — create, write, and read
office-cli ppt create deck.pptx --title "Q2 Review" --author "Alice"
office-cli ppt add-slide deck.pptx --title "Roadmap" --bullets '["Q1: done","Q2: in progress"]' --output deck.pptx
office-cli ppt set-notes deck.pptx --slide 1 --notes "Speaker notes here" --output deck.pptx
office-cli ppt delete-slide deck.pptx --slide 3 --output deck.pptx
office-cli ppt reorder deck.pptx --order "3,1,2" --output deck.pptx
office-cli ppt outline deck.pptx --json
office-cli ppt replace deck.pptx --pairs @placeholders.json
office-cli ppt build deck.pptx --spec '{"title":"Q2 Review","author":"Alice","slides":[{"title":"Overview","bullets":["Revenue up 20%"]}]}'
office-cli ppt build deck.pptx --template company.pptx --spec '{"title":"Q2 Report","slides":[{"title":"Overview"},{"title":"Details"}]}'
office-cli ppt build deck.pptx --spec @deck-spec.json          # @file for large specs

# PDF — 18 commands for full lifecycle
office-cli pdf info report.pdf --json
office-cli pdf merge a.pdf b.pdf --output combined.pdf
office-cli pdf trim big.pdf --pages "1,3,5-7" --output excerpt.pdf
office-cli pdf rotate report.pdf --degrees 90 --pages 1 --output rotated.pdf
office-cli pdf watermark report.pdf --text "DRAFT" --output report.draft.pdf
office-cli pdf encrypt report.pdf --user-password 1234 --output secret.pdf
office-cli pdf optimize report.pdf --output smaller.pdf
office-cli pdf reorder report.pdf --order "3,1,2,4" --output reordered.pdf
office-cli pdf insert-blank report.pdf --after 2 --count 1 --output report.with-blanks.pdf
office-cli pdf bookmarks report.pdf --json
office-cli pdf set-meta report.pdf --title "Annual Report" --author "Alice" --output report.meta.pdf
office-cli pdf stamp-image report.pdf --image logo.png --pages "1-" --output report.stamped.pdf
```

## Architecture

```
office-cli/
├── cmd/
│   └── office-cli/
│       └── main.go               # entrypoint
│   ├── root.go                   # global flags, exit codes, audit hook
│   ├── setup.go                  # config file creation
│   ├── doctor.go                 # environment / tool diagnostics
│   ├── reference.go              # all commands as structured Markdown
│   ├── install_skill.go          # copy skill files to ~/.agent/skills/
│   ├── universal.go              # info / meta / extract-text (auto-detect format)
│   ├── helpers.go                # resolveInput / resolveOutput / path validation
│   ├── excel.go                  # 45 commands: read/write/style/sort/chart/...
│   ├── word.go                   # 25 commands: read/replace/create/add-*/style
│   ├── ppt.go                    # 17 commands: read/replace/create/add-*/build/layout
│   ├── ppt_write.go              # PPT write command implementations
│   ├── pdf.go                    # 22 commands: read/merge/split/bookmarks/...
│   ├── e2e_test.go               # E2E tests (original 44 commands)
│   └── e2e_new_test.go           # E2E tests (36 new commands)
├── internal/
│   ├── output/
│   │   ├── output.go             # color-aware tables, JSON printer
│   │   ├── errors.go             # structured error codes (errorCode + hint)
│   │   └── flat.go               # FlatSheet, FlatPage, FlatSlide, FlatParagraph
│   ├── audit/
│   │   └── audit.go              # JSONL write-command logger
│   ├── config/
│   │   └── config.go             # ~/.office-cli path resolution
│   └── engine/
│       ├── common/
│       │   └── zipxml.go         # RewriteEntries, WriteNewZip, ReadEntry, CoreProps, AppProps
│       ├── excel/
│       │   └── excel.go          # excelize/v2 — 30 functions + 4 types
│       ├── word/
│       │   └── word.go           # native ZIP+XML docx — 8 write functions
│       ├── ppt/
│       │   └── ppt.go            # native ZIP+XML pptx — 8 write functions
│       └── pdf/
│           └── pdf.go            # pdfcpu + ledongthuc/pdf — 6 new functions
├── skills/
│   └── office-cli/
│       ├── SKILL.md              # AI Agent entry point (router + conventions)
│       └── docs/
│           ├── excel.md          # 45 Excel commands reference
│           ├── word.md           # 25 Word commands reference
│           ├── ppt.md            # 17 PPT commands reference
│           └── pdf.md            # 22 PDF commands reference
├── scripts/
│   ├── install.js                # npm postinstall — download binary from GitHub Release
│   ├── run.js                    # npm bin wrapper — exec binary with args
│   └── gen-fixtures/             # test fixture generator (xlsx/docx/pptx/pdf)
├── .github/
│   └── workflows/
│       ├── ci.yml                # test matrix: 3 OS × Go version
│       └── release.yml           # goreleaser + npm publish on v* tag
├── package.json                  # npm package (@fatecannotbealtered-/office-cli)
├── go.mod / go.sum               # Go module dependencies
├── Makefile                      # build / test / vet / fmt / snapshot
├── CHANGELOG.md
├── CONTRIBUTING.md
├── SECURITY.md
├── LICENSE                       # MIT
└── .goreleaser.yml               # cross-compile + archive + checksum config
```

### Why pure Go (no pandoc / libreoffice required)

- Predictable behaviour across machines (no version skew of external tools).
- Smaller, auditable dependency surface.
- The four core formats are well-served by mature Go libraries:
  - **Excel**: `excelize/v2` is the de-facto standard.
  - **PDF**: `pdfcpu` covers structural ops; `ledongthuc/pdf` covers text extraction.
  - **Word/PPT**: docx and pptx are simply ZIP archives of XML — we read/edit them directly.

`office-cli doctor` reports whether `pandoc` or `libreoffice` are installed; they unlock advanced format conversions but are NOT required for any of the commands documented above.

## Design principles

1. **AI-first surface**: every command supports `--json`; errors carry stable codes + hints; reading commands support range/keyword/limit filters to control token usage.
2. **Files are inputs**: the first positional argument is always a file path. Outputs are explicit (`--output`).
3. **Immutable by default**: PDF write commands (`trim`, `merge`, `split`, `watermark`, `reorder`, `insert-blank`, etc.) require `--output` and never mutate the input file. Word/PPT write commands default to in-place mutation for iterative editing convenience; pass `--output` to write to a new file. `excel write` / `excel append` and Excel formatting commands also mutate in place (Excel is the natural exception).
4. **Auditable mutations**: every write command lands in `~/.office-cli/audit/`.
5. **Small dependency surface**: each format engine isolated under `internal/engine/<format>` so swapping libraries is a one-day job.

## Known limitations

### Text replacement

`word replace`, `ppt replace`, and `pdf replace` only match text that lives **entirely within a single XML text element**. This means:

- **Word**: text split across multiple `<w:r>` runs (font boundaries, spell-check splits) will NOT be matched
- **PowerPoint**: text split across multiple `<a:r>` runs will NOT be matched
- **PDF**: only parenthesized string literals in `Tj` operators are matched; hex strings and text split across operators are NOT matched

When replacements produce zero hits, the JSON output includes a `warning` field explaining this limitation.

### Large JSON specs

For `--spec` and `--pairs` flags that accept JSON payloads, prefer the `@file` syntax over inline JSON to avoid shell escaping issues:

```bash
office-cli ppt build deck.pptx --spec @deck-spec.json
office-cli pdf add-text report.pdf --spec @overlays.json --output out.pdf
```

## Permission system

office-cli defaults to `read-only`. Write and destructive commands require elevated permissions set in `~/.office-cli/config.json`:

```json
{ "permissions": { "mode": "write" } }
```

| Level | What's allowed |
|---|---|
| `read-only` (default) | All read commands |
| `write` | + write cells, append rows, replace text, style/sort/freeze/merge cells, add charts/images, create Word/PPT documents, add slides/paragraphs/tables, merge/split/trim/reorder/watermark/stamp PDFs, set PDF metadata/bookmarks |
| `full` | + delete sheets, encrypt/decrypt PDFs |

Override via env var: `OFFICE_CLI_PERMISSIONS=write`. Commands that exceed the configured level exit with code 5.

## Environment

| Variable | Default | Effect |
|---|---|---|
| `OFFICE_CLI_HOME` | `~/.office-cli` | Override home directory |
| `OFFICE_CLI_PERMISSIONS` | `read-only` | Override permission level (`read-only`, `write`, `full`) |
| `OFFICE_CLI_NO_AUDIT` | unset | Set `1` to disable audit logging |
| `OFFICE_CLI_AUDIT_RETENTION_MONTHS` | `3` | Auto-clean audit files older than N months (`0` = keep forever) |
| `NO_COLOR` | unset | Set anything to disable colored output |

## Development

```bash
make build       # build into ./bin/office-cli
make test        # go test ./...
make vet         # go vet
make fmt         # check gofmt cleanliness
make snapshot    # local goreleaser dry run
```

## License

MIT — see `LICENSE`.
