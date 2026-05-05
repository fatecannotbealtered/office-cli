#!/usr/bin/env bash
#
# office-cli end-to-end smoke test (Linux / macOS).
#
# This script:
#   1. Generates fresh fixtures via scripts/gen-fixtures
#   2. Runs every read/write subcommand against those fixtures
#   3. Validates the JSON output for the ones AI Agents most often consume
#
# Override $OFFICE_CLI_BIN to test a prebuilt binary, otherwise we build from source.

set -u
cd "$(dirname "$0")/.."
REPO_ROOT="$(pwd)"

# ---- Build (or reuse) the binary ----------------------------------------------
BIN="${OFFICE_CLI_BIN:-}"
if [[ -z "$BIN" ]]; then
  echo "Building office-cli ..."
  go build -o ./bin/office-cli ./cmd/office-cli || { echo "go build failed" >&2; exit 1; }
  BIN="./bin/office-cli"
fi
BIN="$(cd "$(dirname "$BIN")" && pwd)/$(basename "$BIN")"
echo "Using binary: $BIN"

# ---- Permissions ---------------------------------------------------------------
# E2E tests exercise write commands; grant write permission globally.
export OFFICE_CLI_PERMISSIONS=write

# ---- Workspace -----------------------------------------------------------------
WORK="$REPO_ROOT/tmp/e2e"
rm -rf "$WORK"
mkdir -p "$WORK/fixtures" "$WORK/out"
FIXTURES="$WORK/fixtures"
OUTPUTS="$WORK/out"

go run ./scripts/gen-fixtures --out "$FIXTURES" || { echo "gen-fixtures failed" >&2; exit 1; }

# ---- Test helpers --------------------------------------------------------------
TESTS=0
PASSED=0
FAILED=()

step() {
  local name="$1"; shift
  TESTS=$((TESTS + 1))
  if "$@"; then
    echo "  PASS $name"
    PASSED=$((PASSED + 1))
  else
    echo "  FAIL $name"
    FAILED+=("$name")
  fi
}

# Run the binary, capturing both streams; echo combined output and return its
# exit code. Tests pipe this into jq / grep for validation.
run_bin() {
  "$BIN" "$@" 2>&1
  return ${PIPESTATUS[0]}
}

must_ok() {
  local out
  out="$(run_bin "$@")"
  local code=$?
  if [[ $code -ne 0 ]]; then
    echo "$out" >&2
    return $code
  fi
  printf '%s' "$out"
}

# Convenience: run command with --json and pipe into jq for inspection.
must_jq() {
  local jq_filter="$1"; shift
  must_ok "$@" | jq -e "$jq_filter" >/dev/null
}

# ---- Test definitions ----------------------------------------------------------
DOCX="$FIXTURES/sample.docx"
XLSX="$FIXTURES/sample.xlsx"
XLSX_CJK="$FIXTURES/样本.xlsx"
PPTX="$FIXTURES/sample.pptx"
PDF1="$FIXTURES/sample.pdf"
PDF2="$FIXTURES/sample-2.pdf"
CSV="$FIXTURES/sample.csv"

echo
echo "=== Universal commands ==="
step "info docx" must_jq '.format == "docx"' info "$DOCX" --json
step "info xlsx" must_jq '.format == "xlsx"' info "$XLSX" --json
step "info pptx" must_jq '.format == "pptx"' info "$PPTX" --json
step "info pdf"  must_jq '.format == "pdf"'  info "$PDF1" --json
step "meta xlsx CJK" must_jq '.sheets | index("中文表")' meta "$XLSX_CJK" --json
step "extract-text docx" must_jq '.text | contains("Introduction")' extract-text "$DOCX" --json

echo
echo "=== Excel ==="
step "excel sheets" must_jq 'length >= 2' excel sheets "$XLSX" --json
step "excel read"   must_jq '.rows | length >= 4' excel read "$XLSX" --sheet Sales --json
step "excel cell"   must_jq '.value == "100"' excel cell "$XLSX" B2 --sheet Sales --json
step "excel search" must_jq '.matches | length >= 1' excel search "$XLSX" Bob --json

step "excel append" bash -c '
  copy="$1/appended.xlsx"
  cp "$2" "$copy"
  rows="$1/rows.json"
  printf "%s" "[[\"Dave\",400]]" > "$rows"
  "$3" excel append "$copy" --sheet Sales --rows "@$rows" --force --json >/dev/null || exit 1
  "$3" excel read "$copy" --sheet Sales --json | jq -e ".rows[][0] | select(. == \"Dave\")" >/dev/null
' bash "$OUTPUTS" "$XLSX" "$BIN"

step "excel to-csv (BOM)" bash -c '
  out="$1/out.csv"
  "$2" excel to-csv "$3" --sheet Sales --output "$out" --bom --force --json >/dev/null || exit 1
  head -c 3 "$out" | od -An -tx1 | tr -d " \n" | grep -q "^efbbbf$"
' bash "$OUTPUTS" "$BIN" "$XLSX"

step "excel from-csv" bash -c '
  out="$1/from-csv.xlsx"
  "$2" excel from-csv "$3" --output "$out" --sheet csv --force --json >/dev/null || exit 1
  "$2" excel sheets "$out" --json | jq -e "length >= 1" >/dev/null
' bash "$OUTPUTS" "$BIN" "$CSV"

echo
echo "=== Word ==="
step "word read markdown" must_jq '.markdown | contains("Introduction")' word read "$DOCX" --json --format markdown
step "word stats"         must_jq '.stats.paragraphs >= 5' word stats "$DOCX" --json
step "word headings"      must_jq '.headings | length >= 1' word headings "$DOCX" --json

step "word replace" bash -c '
  out="$1/replaced.docx"
  "$2" word replace "$3" --output "$out" --find "{{NAME}}" --replace "Alice" --force --json >/dev/null || exit 1
  "$2" word read "$out" --json | jq -e "[.paragraphs[].text] | join(\" \") | contains(\"Alice\")" >/dev/null
' bash "$OUTPUTS" "$BIN" "$DOCX"

step "word replace --pairs" bash -c '
  out="$1/replaced-pairs.docx"
  pairs="$1/pairs.json"
  cat > "$pairs" <<EOF
[{"find":"{{NAME}}","replace":"Bob"},{"find":"{{INVOICE}}","replace":"INV-7"}]
EOF
  "$2" word replace "$3" --output "$out" --pairs "@$pairs" --force --json >/dev/null || exit 1
  "$2" word read "$out" --json | jq -e "[.paragraphs[].text] | join(\" \") | (contains(\"Bob\") and contains(\"INV-7\"))" >/dev/null
' bash "$OUTPUTS" "$BIN" "$DOCX"

echo
echo "=== PPT ==="
step "ppt count"   must_jq '.slides == 3' ppt count "$PPTX" --json
step "ppt outline" must_jq '.outline | length == 3' ppt outline "$PPTX" --json
step "ppt replace" bash -c '
  out="$1/replaced.pptx"
  "$2" ppt replace "$3" --output "$out" --find "{{MILESTONE}}" --replace "2026Q3" --force --json >/dev/null || exit 1
  "$2" extract-text "$out" --json | jq -e ".text | contains(\"2026Q3\")" >/dev/null
' bash "$OUTPUTS" "$BIN" "$PPTX"

echo
echo "=== PDF ==="
step "pdf info" must_jq '.pageCount == 1' pdf info "$PDF1" --json
step "pdf merge (positional)" bash -c '
  out="$1/merged.pdf"
  "$2" pdf merge "$3" "$4" --output "$out" --force --json >/dev/null || exit 1
  "$2" pdf info "$out" --json | jq -e ".pageCount == 3" >/dev/null
' bash "$OUTPUTS" "$BIN" "$PDF1" "$PDF2"

step "pdf split" bash -c '
  out="$1/split"
  mkdir -p "$out"
  "$2" pdf split "$3" --output-dir "$out" --span 1 --force --json >/dev/null || exit 1
  ls "$out"/*.pdf | wc -l | awk "{exit !(\$1 >= 2)}"
' bash "$OUTPUTS" "$BIN" "$PDF2"

step "pdf trim" bash -c '
  out="$1/trimmed.pdf"
  "$2" pdf trim "$3" --output "$out" --pages 1 --force --json >/dev/null || exit 1
  "$2" pdf info "$out" --json | jq -e ".pageCount == 1" >/dev/null
' bash "$OUTPUTS" "$BIN" "$PDF2"

echo
echo "=== Doctor / Reference ==="
step "doctor --json" must_jq '.os != null' doctor --json
step "reference"     bash -c '"$1" reference | grep -qi excel' bash "$BIN"

# ---- Summary -------------------------------------------------------------------
echo
echo "========================================="
echo "Tests: $TESTS  Passed: $PASSED  Failed: ${#FAILED[@]}"
if [[ ${#FAILED[@]} -gt 0 ]]; then
  echo "Failures:"
  printf '  - %s\n' "${FAILED[@]}"
  exit 1
fi
exit 0
