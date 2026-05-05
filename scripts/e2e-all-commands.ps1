# office-cli end-to-end smoke test (Windows / PowerShell)
#
# This script:
#   1. Generates fresh fixtures via scripts/gen-fixtures
#   2. Runs every read/write subcommand against those fixtures
#   3. Validates the JSON output for the ones AI Agents most often consume
#
# Override $env:OFFICE_CLI_BIN to test a prebuilt binary, otherwise we build
# from source.

$ErrorActionPreference = "Stop"
$repoRoot = Resolve-Path "$PSScriptRoot/.."
Set-Location $repoRoot

# ---- Build (or reuse) the binary ----------------------------------------------
$bin = $env:OFFICE_CLI_BIN
if (-not $bin) {
  Write-Host "Building office-cli ..." -ForegroundColor Cyan
  go build -o "bin/office-cli.exe" "./cmd/office-cli"
  if ($LASTEXITCODE -ne 0) { throw "go build failed" }
  $bin = "./bin/office-cli.exe"
}
$bin = (Resolve-Path $bin).Path
Write-Host "Using binary: $bin" -ForegroundColor Cyan

# ---- Permissions ---------------------------------------------------------------
# E2E tests exercise write commands; grant write permission globally.
$env:OFFICE_CLI_PERMISSIONS = "write"

# ---- Workspace -----------------------------------------------------------------
$work = Join-Path $repoRoot "tmp/e2e"
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Force -Path $work | Out-Null
$fixtures = Join-Path $work "fixtures"
$outputs = Join-Path $work "out"
New-Item -ItemType Directory -Force -Path $fixtures | Out-Null
New-Item -ItemType Directory -Force -Path $outputs | Out-Null

go run "./scripts/gen-fixtures" --out $fixtures
if ($LASTEXITCODE -ne 0) { throw "gen-fixtures failed" }

# ---- Test helper ---------------------------------------------------------------
$global:tests = 0
$global:passed = 0
$global:failed = @()

function Step {
  param(
    [string]$Name,
    [scriptblock]$Body
  )
  $global:tests++
  try {
    & $Body
    Write-Host "  PASS $Name" -ForegroundColor Green
    $global:passed++
  } catch {
    Write-Host "  FAIL $Name" -ForegroundColor Red
    Write-Host "    $($_.Exception.Message)" -ForegroundColor Red
    $global:failed += $Name
  }
}

function RunBin {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$ArgsList)
  # PowerShell wraps native stderr as ErrorRecord objects; Out-String force-flattens
  # them into a single string and avoids "RemoteException" tokens leaking through.
  $previousEAP = $ErrorActionPreference
  $ErrorActionPreference = "Continue"
  try {
    $out = & $bin @ArgsList 2>&1 | Out-String
  } finally {
    $ErrorActionPreference = $previousEAP
  }
  $code = $LASTEXITCODE
  return @{ Stdout = $out; Code = $code }
}

function MustOk {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$ArgsList)
  $r = RunBin @ArgsList
  if ($r.Code -ne 0) {
    $preview = if ($r.Stdout.Length -gt 500) { $r.Stdout.Substring(0, 500) + "..." } else { $r.Stdout }
    throw "expected exit 0, got $($r.Code): $preview"
  }
  # Always return as a single string so callers can use String.Contains / -match.
  return [string]$r.Stdout
}

function MustJSON {
  param([Parameter(ValueFromRemainingArguments=$true)][string[]]$ArgsList)
  $stdout = MustOk @ArgsList
  return $stdout | ConvertFrom-Json
}

# ---- Tests ---------------------------------------------------------------------
$docx = Join-Path $fixtures "sample.docx"
$xlsx = Join-Path $fixtures "sample.xlsx"
$xlsxCJK = Join-Path $fixtures "样本.xlsx"
$pptx = Join-Path $fixtures "sample.pptx"
$pdf1 = Join-Path $fixtures "sample.pdf"
$pdf2 = Join-Path $fixtures "sample-2.pdf"
$csv = Join-Path $fixtures "sample.csv"

Write-Host ""
Write-Host "=== Universal commands ===" -ForegroundColor Yellow
Step "info docx" { $j = MustJSON info $docx --json; if ($j.format -ne "docx") { throw "format mismatch" } }
Step "info xlsx" { $j = MustJSON info $xlsx --json; if ($j.format -ne "xlsx") { throw "format mismatch" } }
Step "info pptx" { $j = MustJSON info $pptx --json; if ($j.format -ne "pptx") { throw "format mismatch" } }
Step "info pdf"  { $j = MustJSON info $pdf1 --json; if ($j.format -ne "pdf") { throw "format mismatch" } }
Step "meta xlsx CJK path" {
  $j = MustJSON meta $xlsxCJK --json
  if (-not ($j.sheets -contains "中文表")) { throw "expected CJK sheet" }
}
Step "extract-text docx" {
  $j = MustJSON extract-text $docx --json
  if (-not $j.text.Contains("Introduction")) { throw "missing text" }
}

Write-Host ""
Write-Host "=== Excel ===" -ForegroundColor Yellow
Step "excel sheets" { $j = MustJSON excel sheets $xlsx --json; if ($j.Count -lt 2) { throw "expected at least 2 sheets" } }
Step "excel read" { $j = MustJSON excel read $xlsx --sheet Sales --json; if ($j.rows.Count -lt 4) { throw "expected at least 4 rows" } }
Step "excel cell" { $j = MustJSON excel cell $xlsx B2 --sheet Sales --json; if ($j.value -ne "100") { throw "expected 100, got $($j.value)" } }
Step "excel search" { $j = MustJSON excel search $xlsx Bob --json; if ($j.matches.Count -lt 1) { throw "expected at least 1 match" } }
Step "excel append" {
  $copy = Join-Path $outputs "appended.xlsx"
  Copy-Item $xlsx $copy
  # PowerShell strips double-quotes when forwarding to native exes; use @file.json.
  $rowsFile = Join-Path $outputs "rows.json"
  Set-Content -Path $rowsFile -Value '[["Dave",400]]' -Encoding UTF8
  MustOk excel append $copy --sheet Sales --rows "@$rowsFile" --force | Out-Null
  $j = MustJSON excel read $copy --sheet Sales --json
  $names = $j.rows | ForEach-Object { $_[0] }
  if (-not ($names -contains "Dave")) { throw "Dave not appended" }
}
Step "excel to-csv (with BOM)" {
  $out = Join-Path $outputs "out.csv"
  MustOk excel to-csv $xlsx --sheet Sales --output $out --bom --force | Out-Null
  $bytes = [System.IO.File]::ReadAllBytes($out)
  if ($bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) { throw "missing BOM" }
}
Step "excel from-csv" {
  $out = Join-Path $outputs "from-csv.xlsx"
  MustOk excel from-csv $csv --output $out --sheet csv --force | Out-Null
  $j = MustJSON excel sheets $out --json
  if ($j.Count -lt 1) { throw "no sheets created" }
}

Write-Host ""
Write-Host "=== Word ===" -ForegroundColor Yellow
Step "word read markdown" {
  $j = MustJSON word read $docx --json --format markdown
  if (-not $j.markdown.Contains("Introduction")) { throw "missing heading" }
}
Step "word stats" {
  $j = MustJSON word stats $docx --json
  if ($j.stats.paragraphs -lt 5) { throw "expected at least 5 paragraphs" }
}
Step "word headings" {
  $j = MustJSON word headings $docx --json
  if ($j.headings.Count -lt 1) { throw "expected at least 1 heading" }
}
Step "word replace" {
  $out = Join-Path $outputs "replaced.docx"
  MustOk word replace $docx --output $out --find "{{NAME}}" --replace "Alice" --force | Out-Null
  $j = MustJSON word read $out --json
  $joined = ($j.paragraphs | ForEach-Object { $_.text }) -join " "
  if (-not $joined.Contains("Alice")) { throw "replacement not applied" }
}
Step "word replace --pairs" {
  $out = Join-Path $outputs "replaced-pairs.docx"
  $pairsFile = Join-Path $outputs "pairs.json"
  Set-Content -Path $pairsFile -Value '[{"find":"{{NAME}}","replace":"Bob"},{"find":"{{INVOICE}}","replace":"INV-7"}]' -Encoding UTF8
  MustOk word replace $docx --output $out --pairs "@$pairsFile" --force | Out-Null
  $j = MustJSON word read $out --json
  $joined = ($j.paragraphs | ForEach-Object { $_.text }) -join " "
  if (-not ($joined.Contains("Bob") -and $joined.Contains("INV-7"))) { throw "batch replacement failed" }
}

Write-Host ""
Write-Host "=== PPT ===" -ForegroundColor Yellow
Step "ppt count" { $j = MustJSON ppt count $pptx --json; if ($j.slides -ne 3) { throw "expected 3" } }
Step "ppt outline" {
  $j = MustJSON ppt outline $pptx --json
  if ($j.outline.Count -ne 3) { throw "expected 3 entries" }
}
Step "ppt replace" {
  $out = Join-Path $outputs "replaced.pptx"
  MustOk ppt replace $pptx --output $out --find "{{MILESTONE}}" --replace "2026Q3" --force | Out-Null
  $j = MustJSON extract-text $out --json
  if (-not $j.text.Contains("2026Q3")) { throw "ppt replacement not applied" }
}

Write-Host ""
Write-Host "=== PDF ===" -ForegroundColor Yellow
Step "pdf info" { $j = MustJSON pdf info $pdf1 --json; if ($j.pageCount -ne 1) { throw "expected 1 page" } }
Step "pdf merge (positional)" {
  $out = Join-Path $outputs "merged.pdf"
  MustOk pdf merge $pdf1 $pdf2 --output $out --force | Out-Null
  $j = MustJSON pdf info $out --json
  if ($j.pageCount -ne 3) { throw "expected 3 pages, got $($j.pageCount)" }
}
Step "pdf split" {
  $out = Join-Path $outputs "split"
  New-Item -ItemType Directory -Force -Path $out | Out-Null
  MustOk pdf split $pdf2 --output-dir $out --span 1 --force | Out-Null
  $files = Get-ChildItem -Path $out -Filter *.pdf
  if ($files.Count -lt 2) { throw "expected at least 2 split files" }
}
Step "pdf trim" {
  $out = Join-Path $outputs "trimmed.pdf"
  MustOk pdf trim $pdf2 --output $out --pages 1 --force | Out-Null
  $j = MustJSON pdf info $out --json
  if ($j.pageCount -ne 1) { throw "expected 1 page" }
}

Write-Host ""
Write-Host "=== Doctor / Reference ===" -ForegroundColor Yellow
Step "doctor --json" {
  $j = MustJSON doctor --json
  if (-not $j.os) { throw "missing os field" }
}
Step "reference" {
  $stdout = MustOk reference
  $stdoutStr = if ($stdout -is [string]) { $stdout } else { $stdout -join "`n" }
  if ($stdoutStr -notmatch "(?i)excel") { throw "missing excel section" }
}

# ---- Summary -------------------------------------------------------------------
Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Tests: $($global:tests)  Passed: $($global:passed)  Failed: $($global:failed.Count)" -ForegroundColor Cyan
if ($global:failed.Count -gt 0) {
  Write-Host "Failures:" -ForegroundColor Red
  $global:failed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
  exit 1
}
exit 0
