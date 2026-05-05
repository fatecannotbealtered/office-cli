package cmd_test

import (
	"bytes"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

// binPath is set by TestMain after building the binary.
var binPath string

// fixtureDir is set by TestMain after generating fixtures.
var fixtureDir string

func TestMain(m *testing.M) {
	dir, err := os.MkdirTemp("", "office-cli-e2e-*")
	if err != nil {
		panic(err)
	}
	defer func() { _ = os.RemoveAll(dir) }()

	name := "office-cli"
	if runtime.GOOS == "windows" {
		name += ".exe"
	}
	binPath = filepath.Join(dir, name)

	// Build the binary.
	modRoot := findModuleRoot()
	build := exec.Command("go", "build", "-o", binPath, "./cmd/office-cli")
	build.Dir = modRoot
	build.Stdout = os.Stdout
	build.Stderr = os.Stderr
	if err := build.Run(); err != nil {
		panic("build failed: " + err.Error())
	}

	// Generate fixtures using the existing generator.
	fixtureDir = filepath.Join(dir, "fixtures")
	gen := exec.Command("go", "run", "./scripts/gen-fixtures", "--out", fixtureDir)
	gen.Dir = modRoot
	gen.Stdout = os.Stdout
	gen.Stderr = os.Stderr
	if err := gen.Run(); err != nil {
		panic("gen-fixtures failed: " + err.Error())
	}

	os.Exit(m.Run())
}

func findModuleRoot() string {
	dir, _ := os.Getwd()
	for {
		if _, err := os.Stat(filepath.Join(dir, "go.mod")); err == nil {
			return dir
		}
		parent := filepath.Dir(dir)
		if parent == dir {
			break
		}
		dir = parent
	}
	return "."
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

type result struct {
	Stdout   string
	Stderr   string
	ExitCode int
}

func run(t *testing.T, env []string, args ...string) result {
	t.Helper()
	cmd := exec.Command(binPath, args...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if env != nil {
		cmd.Env = append(os.Environ(), env...)
	}
	err := cmd.Run()
	exitCode := 0
	if err != nil {
		if exitErr, ok := err.(*exec.ExitError); ok {
			exitCode = exitErr.ExitCode()
		} else {
			t.Fatalf("exec failed: %v", err)
		}
	}
	return result{Stdout: stdout.String(), Stderr: stderr.String(), ExitCode: exitCode}
}

// setupHome creates a temp home dir, runs setup, returns the env slice.
func setupHome(t *testing.T, perm string) []string {
	t.Helper()
	home := t.TempDir()
	env := []string{"OFFICE_CLI_HOME=" + home}
	if perm != "" {
		env = append(env, "OFFICE_CLI_PERMISSIONS="+perm)
	}
	r := run(t, env, "setup", "--json")
	if r.ExitCode != 0 {
		t.Fatalf("setup failed: exit=%d stderr=%s", r.ExitCode, r.Stderr)
	}
	return env
}

func parseJSON(t *testing.T, s string, v any) {
	t.Helper()
	s = strings.TrimSpace(s)
	if s == "" {
		t.Fatal("empty JSON string")
	}
	if err := json.Unmarshal([]byte(s), v); err != nil {
		t.Fatalf("invalid JSON: %v\nraw: %s", err, s[:min(len(s), 500)])
	}
}

func requireField(t *testing.T, m map[string]any, key string) {
	t.Helper()
	if _, ok := m[key]; !ok {
		t.Errorf("missing required field %q in %v", key, m)
	}
}

func fixture(name string) string {
	return filepath.Join(fixtureDir, name)
}

// ============================================================================
// GLOBAL FLAGS
// ============================================================================

func TestHelp(t *testing.T) {
	r := run(t, nil, "--help")
	if r.ExitCode != 0 {
		t.Fatalf("exit=%d", r.ExitCode)
	}
	if !strings.Contains(r.Stdout, "excel") || !strings.Contains(r.Stdout, "word") {
		t.Error("help should list excel and word subcommands")
	}
}

func TestVersion(t *testing.T) {
	r := run(t, nil, "--version")
	if r.ExitCode != 0 {
		t.Fatalf("exit=%d", r.ExitCode)
	}
	if strings.TrimSpace(r.Stdout) == "" {
		t.Error("version output is empty")
	}
}

// ============================================================================
// SETUP
// ============================================================================

func TestSetup_CreatesConfig(t *testing.T) {
	home := t.TempDir()
	r := run(t, []string{"OFFICE_CLI_HOME=" + home}, "setup", "--json")
	if r.ExitCode != 0 {
		t.Fatalf("exit=%d stderr=%s", r.ExitCode, r.Stderr)
	}
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "created" {
		t.Errorf("status=%v want created", m["status"])
	}
	requireField(t, m, "path")
	requireField(t, m, "mode")
}

func TestSetup_Idempotent(t *testing.T) {
	home := t.TempDir()
	env := []string{"OFFICE_CLI_HOME=" + home}
	run(t, env, "setup", "--json")
	r := run(t, env, "setup", "--json")
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("second run status=%v want ok", m["status"])
	}
}

// ============================================================================
// PERMISSION SYSTEM
// ============================================================================

func TestPerm_WriteDeniedOnReadOnly(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "write", "dummy.xlsx", "--ref", "A1", "--value", "x", "--json")
	assertDenied(t, r)
}

func TestPerm_ReadAllowedOnReadOnly(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "sheets", "nonexistent.xlsx", "--json")
	if r.ExitCode == 5 {
		t.Errorf("read command should not be permission-denied")
	}
}

func TestPerm_EnvOverride(t *testing.T) {
	env := setupHome(t, "read-only")
	env = append(env, "OFFICE_CLI_PERMISSIONS=write")
	r := run(t, env, "excel", "write", "dummy.xlsx", "--ref", "A1", "--value", "x", "--json")
	if r.ExitCode == 5 {
		t.Errorf("env override should grant write, got permission-denied")
	}
}

func TestPerm_DeleteSheetNeedsFull(t *testing.T) {
	env := setupHome(t, "write")
	r := run(t, env, "excel", "delete-sheet", "dummy.xlsx", "--sheet", "S", "--force", "--json")
	assertDenied(t, r)
}

func TestPerm_EncryptNeedsFull(t *testing.T) {
	env := setupHome(t, "write")
	r := run(t, env, "pdf", "encrypt", "dummy.pdf", "--user-password", "x", "--output", "o.pdf", "--json")
	assertDenied(t, r)
}

func assertDenied(t *testing.T, r result) {
	t.Helper()
	if r.ExitCode != 5 {
		t.Errorf("exit=%d want 5 (FORBIDDEN)\nstderr: %s", r.ExitCode, r.Stderr)
	}
	if !strings.Contains(r.Stderr, "PERMISSION_DENIED") {
		t.Errorf("stderr should contain PERMISSION_DENIED")
	}
}

// ============================================================================
// ERROR FORMAT (AI-friendliness check)
// ============================================================================

func TestError_FileNotFound_HasHint(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "sheets", "/no/such/file.xlsx", "--json")
	assertExit(t, r, 4)
	assertJSONError(t, r.Stderr, "FILE_NOT_FOUND", true)
}

func TestError_InvalidFormat_HasHint(t *testing.T) {
	env := setupHome(t, "read-only")
	txt := filepath.Join(t.TempDir(), "bad.txt")
	_ = os.WriteFile(txt, []byte("x"), 0600)
	r := run(t, env, "excel", "sheets", txt, "--json")
	assertExit(t, r, 2)
	assertJSONError(t, r.Stderr, "INVALID_FORMAT", true)
}

func TestError_MissingArgs(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", "--json")
	if r.ExitCode == 0 {
		t.Error("missing arg should fail")
	}
}

func assertExit(t *testing.T, r result, want int) {
	t.Helper()
	if r.ExitCode != want {
		t.Errorf("exit=%d want %d\nstdout: %s\nstderr: %s", r.ExitCode, want, r.Stdout, r.Stderr)
	}
}

func assertJSONError(t *testing.T, stderr, wantCode string, wantHint bool) {
	t.Helper()
	var m map[string]any
	parseJSON(t, stderr, &m)
	if m["errorCode"] != wantCode {
		t.Errorf("errorCode=%v want %v", m["errorCode"], wantCode)
	}
	if wantHint {
		hint, _ := m["hint"].(string)
		if hint == "" {
			t.Error("hint should be present for AI Agent corrective action")
		}
	}
}

// ============================================================================
// DOCTOR
// ============================================================================

func TestDoctor(t *testing.T) {
	r := run(t, nil, "doctor", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "version")
	requireField(t, m, "goVersion")
	requireField(t, m, "os")
	requireField(t, m, "arch")
	requireField(t, m, "tools")
}

// ============================================================================
// REFERENCE
// ============================================================================

func TestReference(t *testing.T) {
	r := run(t, nil, "reference")
	assertExit(t, r, 0)
	for _, want := range []string{"excel", "word", "ppt", "pdf"} {
		if !strings.Contains(r.Stdout, want) {
			t.Errorf("reference should contain %q", want)
		}
	}
}

// ============================================================================
// DRY-RUN
// ============================================================================

func TestDryRun_ExcelWrite(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "dry.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["a"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "write", out, "--ref", "A1", "--value", "CHANGED", "--dry-run", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["dryRun"] != true {
		t.Errorf("dryRun=%v want true", m["dryRun"])
	}
	// Verify file was NOT modified
	r2 := run(t, env, "excel", "cell", out, "A1", "--sheet", "S", "--json")
	var cell map[string]any
	parseJSON(t, r2.Stdout, &cell)
	if cell["value"] == "CHANGED" {
		t.Error("dry-run should NOT modify the file")
	}
}

// ============================================================================
// EXCEL — ALL SUBCOMMANDS
// ============================================================================

func TestExcel_Sheets(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "sheets", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var sheets []map[string]any
	parseJSON(t, r.Stdout, &sheets)
	if len(sheets) < 1 {
		t.Fatal("expected at least 1 sheet")
	}
	// AI-friendly: flat objects with predictable fields
	for _, s := range sheets {
		requireField(t, s, "name")
		requireField(t, s, "rows")
		requireField(t, s, "cols")
	}
}

func TestExcel_Read_FullSheet(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "sheet")
	requireField(t, m, "rows")
}

func TestExcel_Read_Range(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--range", "A1:B3", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "rows")
}

func TestExcel_Read_WithHeaders(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--with-headers", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "records")
}

func TestExcel_Read_SpecificSheet(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--sheet", "Sales", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["sheet"] != "Sales" {
		t.Errorf("sheet=%v want Sales", m["sheet"])
	}
}

func TestExcel_Cell(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "cell", fixture("sample.xlsx"), "A1", "--sheet", "Sales", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "value")
	requireField(t, m, "file")
}

func TestExcel_Search(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "search", fixture("sample.xlsx"), "Alice", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "count")
	requireField(t, m, "matches")
	count, _ := m["count"].(float64)
	if count < 1 {
		t.Errorf("expected >= 1 match, got %v", m["count"])
	}
}

func TestExcel_Info(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "info", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "file")
	requireField(t, m, "sheets")
}

func TestExcel_Create_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "new.xlsx")
	spec := `{"sheets":[{"name":"Data","rows":[["A","B"],[1,2],["x","y"]]}]}`
	r := run(t, env, "excel", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Read back
	r2 := run(t, env, "excel", "read", out, "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	rows, _ := read["rows"].([]any)
	if len(rows) != 3 {
		t.Errorf("expected 3 rows, got %d", len(rows))
	}
}

func TestExcel_Write_CellRoundtrip(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "wr.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["a","b"],["c","d"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "write", out, "--sheet", "S", "--ref", "A1", "--value", "MODIFIED", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["written"].(float64) != 1 {
		t.Errorf("written=%v want 1", m["written"])
	}

	// Read back
	r2 := run(t, env, "excel", "cell", out, "A1", "--sheet", "S", "--json")
	var cell map[string]any
	parseJSON(t, r2.Stdout, &cell)
	if cell["value"] != "MODIFIED" {
		t.Errorf("value=%v want MODIFIED", cell["value"])
	}
}

func TestExcel_Append(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "app.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["name","val"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "append", out, "--sheet", "S", "--rows", `[["Alice",100],["Bob",200]]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "appended")
	requireField(t, m, "startRow")
}

func TestExcel_ToCSV(t *testing.T) {
	env := setupHome(t, "write")
	csvOut := filepath.Join(t.TempDir(), "out.csv")
	r := run(t, env, "excel", "to-csv", fixture("sample.xlsx"), "--sheet", "Sales", "--output", csvOut, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	if _, err := os.Stat(csvOut); err != nil {
		t.Errorf("CSV file not created: %v", err)
	}
}

func TestExcel_FromCSV(t *testing.T) {
	env := setupHome(t, "write")
	xlsxOut := filepath.Join(t.TempDir(), "fromcsv.xlsx")
	r := run(t, env, "excel", "from-csv", fixture("sample.csv"), "--output", xlsxOut, "--sheet", "Imported", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	requireField(t, m, "rows")
}

func TestExcel_RenameSheet(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "ren.xlsx")
	spec := `{"sheets":[{"name":"Old","rows":[["a"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "rename-sheet", out, "--from", "Old", "--to", "New", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
}

func TestExcel_CopySheet(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "copy.xlsx")
	spec := `{"sheets":[{"name":"Template","rows":[["a"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "copy-sheet", out, "--from", "Template", "--to", "Copy", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	requireField(t, m, "index")
}

func TestExcel_DeleteSheet(t *testing.T) {
	env := setupHome(t, "full")
	out := filepath.Join(t.TempDir(), "del.xlsx")
	spec := `{"sheets":[{"name":"Keep","rows":[["a"]]},{"name":"Drop","rows":[["b"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	r := run(t, env, "excel", "delete-sheet", out, "--sheet", "Drop", "--force", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
}

func TestExcel_SheetNotFound(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--sheet", "NONEXISTENT", "--json")
	assertExit(t, r, 4)
	assertJSONError(t, r.Stderr, "NOT_FOUND", true)
}

// ============================================================================
// WORD — ALL SUBCOMMANDS
// ============================================================================

func TestWord_Read_Paragraphs(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "file")
	requireField(t, m, "paragraphs")
	paras, _ := m["paragraphs"].([]any)
	if len(paras) == 0 {
		t.Error("expected paragraphs")
	}
}

func TestWord_Read_Markdown(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--format", "markdown", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "markdown")
}

func TestWord_Read_Text(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--format", "text", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

func TestWord_Read_Keyword(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--keyword", "Introduction", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	paras, _ := m["paragraphs"].([]any)
	if len(paras) == 0 {
		t.Error("keyword filter should match at least 1 paragraph")
	}
}

func TestWord_Read_Limit(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--limit", "2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	paras, _ := m["paragraphs"].([]any)
	if len(paras) > 2 {
		t.Errorf("limit=2 but got %d paragraphs", len(paras))
	}
}

func TestWord_Replace_Single(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "replaced.docx")
	r := run(t, env, "word", "replace", fixture("sample.docx"), "--find", "{{NAME}}", "--replace", "Claude", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	requireField(t, m, "hits")
	// Verify replacement took effect
	r2 := run(t, env, "word", "read", out, "--format", "text", "--json")
	var text map[string]any
	parseJSON(t, r2.Stdout, &text)
	if !strings.Contains(text["text"].(string), "Claude") {
		t.Error("replacement not found in output")
	}
}

func TestWord_Replace_BatchPairs(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "batch.docx")
	pairs := `[{"find":"{{NAME}}","replace":"Agent"},{"find":"{{INVOICE}}","replace":"INV-001"}]`
	r := run(t, env, "word", "replace", fixture("sample.docx"), "--pairs", pairs, "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	hits, _ := m["hits"].(float64)
	if hits < 2 {
		t.Errorf("expected >= 2 hits, got %v", m["hits"])
	}
}

func TestWord_Meta(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "meta", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "path")
	requireField(t, m, "sizeBytes")
}

func TestWord_Stats(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "stats", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "stats")
	stats, _ := m["stats"].(map[string]any)
	requireField(t, stats, "paragraphs")
	requireField(t, stats, "words")
	requireField(t, stats, "characters")
}

func TestWord_Headings(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "headings", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "headings")
}

func TestWord_Images(t *testing.T) {
	env := setupHome(t, "write")
	outDir := t.TempDir()
	r := run(t, env, "word", "images", fixture("sample.docx"), "--output-dir", outDir, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "images")
}

func TestWord_Replace_NoMatch(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "nomatch.docx")
	r := run(t, env, "word", "replace", fixture("sample.docx"), "--find", "ZZZZNOTFOUND", "--replace", "x", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["hits"].(float64) != 0 {
		t.Errorf("hits=%v want 0", m["hits"])
	}
}

// ============================================================================
// PPT — ALL SUBCOMMANDS
// ============================================================================

func TestPPT_Read_Outline(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "read", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "file")
	requireField(t, m, "slides")
	slides, _ := m["slides"].([]any)
	if len(slides) < 1 {
		t.Error("expected slides")
	}
}

func TestPPT_Read_Markdown(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "read", fixture("sample.pptx"), "--format", "markdown", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "markdown")
}

func TestPPT_Read_Text(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "read", fixture("sample.pptx"), "--format", "text", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

func TestPPT_Read_Keyword(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "read", fixture("sample.pptx"), "--keyword", "Roadmap", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	slides, _ := m["slides"].([]any)
	if len(slides) == 0 {
		t.Error("keyword 'Roadmap' should match at least 1 slide")
	}
}

func TestPPT_Read_SingleSlide(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "read", fixture("sample.pptx"), "--slide", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	slides, _ := m["slides"].([]any)
	if len(slides) != 1 {
		t.Errorf("expected 1 slide, got %d", len(slides))
	}
}

func TestPPT_Count(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "count", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "slides")
	n, _ := m["slides"].(float64)
	if n < 1 {
		t.Errorf("slides=%v want >= 1", m["slides"])
	}
}

func TestPPT_Outline(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "outline", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "outline")
}

func TestPPT_Replace(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "replaced.pptx")
	r := run(t, env, "ppt", "replace", fixture("sample.pptx"), "--find", "{{MILESTONE}}", "--replace", "2026Q3", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	requireField(t, m, "hits")
}

func TestPPT_Meta(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "ppt", "meta", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "path")
	requireField(t, m, "sizeBytes")
	requireField(t, m, "slides")
}

func TestPPT_Images(t *testing.T) {
	env := setupHome(t, "write")
	outDir := t.TempDir()
	r := run(t, env, "ppt", "images", fixture("sample.pptx"), "--output-dir", outDir, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "images")
}

// ============================================================================
// PDF — ALL SUBCOMMANDS
// ============================================================================

func TestPDF_Pages(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "pages", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "pages")
}

func TestPDF_Pages_Dimensions(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "pages", fixture("sample.pdf"), "--dimensions", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "dimensions")
}

func TestPDF_Read_All(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "read", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "pages")
}

func TestPDF_Read_TextOnly(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "read", fixture("sample.pdf"), "--text-only", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

func TestPDF_Read_SinglePage(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "read", fixture("sample-2.pdf"), "--page", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	pages, _ := m["pages"].([]any)
	if len(pages) != 1 {
		t.Errorf("expected 1 page, got %d", len(pages))
	}
}

func TestPDF_Info(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "pdf", "info", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "pageCount")
}

func TestPDF_Merge(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "merged.pdf")
	r := run(t, env, "pdf", "merge", fixture("sample.pdf"), fixture("sample-2.pdf"), "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	// Verify merged file has combined pages
	r2 := run(t, env, "pdf", "pages", out, "--json")
	var pages map[string]any
	parseJSON(t, r2.Stdout, &pages)
	if pages["pages"].(float64) != 3 {
		t.Errorf("merged pages=%v want 3", pages["pages"])
	}
}

func TestPDF_Split(t *testing.T) {
	env := setupHome(t, "write")
	outDir := t.TempDir()
	r := run(t, env, "pdf", "split", fixture("sample-2.pdf"), "--span", "1", "--output-dir", outDir, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	files, _ := m["files"].([]any)
	if len(files) != 2 {
		t.Errorf("expected 2 split files, got %d", len(files))
	}
}

func TestPDF_Trim(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "trimmed.pdf")
	r := run(t, env, "pdf", "trim", fixture("sample-2.pdf"), "--pages", "1", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	r2 := run(t, env, "pdf", "pages", out, "--json")
	var pages map[string]any
	parseJSON(t, r2.Stdout, &pages)
	if pages["pages"].(float64) != 1 {
		t.Errorf("trimmed pages=%v want 1", pages["pages"])
	}
}

func TestPDF_Watermark(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "watermarked.pdf")
	r := run(t, env, "pdf", "watermark", fixture("sample.pdf"), "--text", "DRAFT", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
}

func TestPDF_Rotate(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "rotated.pdf")
	r := run(t, env, "pdf", "rotate", fixture("sample.pdf"), "--degrees", "90", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
}

func TestPDF_Optimize(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "opt.pdf")
	r := run(t, env, "pdf", "optimize", fixture("sample.pdf"), "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
	requireField(t, m, "sizeBefore")
	requireField(t, m, "sizeAfter")
}

func TestPDF_Encrypt_Decrypt(t *testing.T) {
	env := setupHome(t, "full")
	encrypted := filepath.Join(t.TempDir(), "enc.pdf")
	decrypted := filepath.Join(t.TempDir(), "dec.pdf")

	// Encrypt
	r := run(t, env, "pdf", "encrypt", fixture("sample.pdf"), "--user-password", "test123", "--output", encrypted, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("encrypt status=%v", m["status"])
	}

	// Decrypt
	r2 := run(t, env, "pdf", "decrypt", encrypted, "--user-password", "test123", "--output", decrypted, "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	if m2["status"] != "ok" {
		t.Errorf("decrypt status=%v", m2["status"])
	}
}

func TestPDF_ExtractImages(t *testing.T) {
	env := setupHome(t, "write")
	outDir := t.TempDir()
	r := run(t, env, "pdf", "extract-images", fixture("sample.pdf"), "--output-dir", outDir, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v", m["status"])
	}
}

// ============================================================================
// UNIVERSAL COMMANDS
// ============================================================================

func TestInfo_XLSX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "info", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "format")
	requireField(t, m, "sizeBytes")
	if m["format"] != "xlsx" {
		t.Errorf("format=%v want xlsx", m["format"])
	}
	requireField(t, m, "sheets")
}

func TestInfo_DOCX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "info", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["format"] != "docx" {
		t.Errorf("format=%v want docx", m["format"])
	}
}

func TestInfo_PPTX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "info", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["format"] != "pptx" {
		t.Errorf("format=%v want pptx", m["format"])
	}
	requireField(t, m, "slides")
}

func TestInfo_PDF(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "info", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["format"] != "pdf" {
		t.Errorf("format=%v want pdf", m["format"])
	}
	requireField(t, m, "pages")
}

func TestMeta_XLSX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "meta", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "format")
	requireField(t, m, "path")
	requireField(t, m, "sizeBytes")
	requireField(t, m, "sheets")
}

func TestMeta_DOCX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "meta", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "title")
	requireField(t, m, "author")
}

func TestMeta_PPTX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "meta", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "slides")
}

func TestMeta_PDF(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "meta", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "format")
	requireField(t, m, "pages")
}

func TestExtractText_XLSX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "extract-text", fixture("sample.xlsx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
	requireField(t, m, "format")
}

func TestExtractText_DOCX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "extract-text", fixture("sample.docx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

func TestExtractText_PPTX(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "extract-text", fixture("sample.pptx"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

func TestExtractText_PDF(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "extract-text", fixture("sample.pdf"), "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

// ============================================================================
// QUIET MODE
// ============================================================================

func TestQuietMode_SuppressesStdout(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "excel", "read", fixture("sample.xlsx"), "--json", "--quiet")
	assertExit(t, r, 0)
	// stdout should be clean JSON only
	var m map[string]any
	parseJSON(t, strings.TrimSpace(r.Stdout), &m)
}

// ============================================================================
// AUDIT LOGGING
// ============================================================================

func TestAudit_WriteCommandLogged(t *testing.T) {
	home := t.TempDir()
	env := []string{"OFFICE_CLI_HOME=" + home, "OFFICE_CLI_PERMISSIONS=write"}
	run(t, env, "setup", "--json")

	out := filepath.Join(t.TempDir(), "audit.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["a"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	// Check audit dir exists and has files
	auditDir := filepath.Join(home, "audit")
	entries, _ := os.ReadDir(auditDir)
	if len(entries) == 0 {
		t.Error("audit log should have been created for write command")
	}
}

func TestAudit_DisabledByEnv(t *testing.T) {
	home := t.TempDir()
	env := []string{"OFFICE_CLI_HOME=" + home, "OFFICE_CLI_PERMISSIONS=write", "OFFICE_CLI_NO_AUDIT=1"}
	run(t, env, "setup", "--json")

	out := filepath.Join(t.TempDir(), "noaudit.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["a"]]}]}`
	run(t, env, "excel", "create", out, "--spec", spec, "--json")

	auditDir := filepath.Join(home, "audit")
	entries, _ := os.ReadDir(auditDir)
	if len(entries) != 0 {
		t.Errorf("audit should be disabled, found %d files", len(entries))
	}
}
