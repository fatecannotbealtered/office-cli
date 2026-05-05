# Contributing to office-cli

Thanks for taking the time to contribute! office-cli aims to be a small, predictable
toolkit that humans and AI Agents can both rely on — please keep that bar in mind.

## Quick start

```bash
git clone https://github.com/fatecannotbealtered/office-cli.git
cd office-cli
make build
./bin/office-cli doctor
```

Run the full test/lint suite before opening a PR:

```bash
make fmt vet test
```

## Project layout

| Path | Purpose |
|---|---|
| `cmd/` | One file per top-level resource (excel, word, ppt, pdf) plus shared scaffolding (root, doctor, reference). |
| `internal/output/` | Color-aware printing, structured JSON errors, table renderer, flat types for token-efficient output. |
| `internal/audit/` | JSONL audit logger for write commands. |
| `internal/config/` | Resolves `~/.office-cli` and friends. |
| `internal/engine/<format>/` | One package per document family (excel, pdf, word, ppt) plus a shared `common` for zipxml helpers. |
| `skills/office-cli/SKILL.md` | The AI Agent skill that ships with the binary. |

## Adding a new command

1. Add the engine logic under `internal/engine/<format>/` if a new operation is needed.
2. Add the cobra subcommand in `cmd/<format>.go`.
3. If the command writes to disk, call `markWrite(cmd)` in `init()` so audit logging picks it up.
4. Wire `--json`, `--quiet`, `--dry-run` and `--force` consistently with the existing commands.
5. Update `skills/office-cli/SKILL.md` with the new command and an example.

## Coding standards

- Stick to standard library + the four chosen engines (`excelize/v2`, `pdfcpu`, `ledongthuc/pdf`, plus our own zip+xml work). Adding a new external dep needs justification.
- Document **every** exported symbol. Public APIs of the engine packages are consumed by the cobra layer and need to be readable on their own.
- Errors must funnel through `emitError` / `emitFileError` so `--json` and `--quiet` continue to work.
- Prefer flat / token-efficient JSON shapes (`output.FlatSheet`, `output.FlatPage`, etc.) over raw API responses.
- Keep human output minimal under `--quiet`; AI Agents should always be able to add `--json --quiet` and parse stdout cleanly.

## Reporting issues

Please include:
- `office-cli doctor --json` output
- The exact command you ran
- For document-specific bugs: a minimal sample file (or instructions to reproduce one)

## License

By contributing you agree your work is licensed under the MIT License (see `LICENSE`).
