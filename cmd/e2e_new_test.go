package cmd_test

import (
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// ============================================================================
// EXCEL NEW COMMANDS
// ============================================================================

func TestExcel_ReadTyped(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "read", out, "--typed", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "headers")
	requireField(t, m, "rows")
	requireField(t, m, "types")

	// Verify types are correct
	types, _ := m["types"].(map[string]any)
	if types["Name"] != "string" {
		t.Errorf("Name type=%v want string", types["Name"])
	}
	if types["Score"] != "number" {
		t.Errorf("Score type=%v want number", types["Score"])
	}

	// Verify values are typed (numbers as numbers, not strings)
	rows, _ := m["rows"].([]any)
	if len(rows) == 0 {
		t.Fatal("expected data rows")
	}
	firstRow, _ := rows[0].(map[string]any)
	if _, ok := firstRow["Score"].(float64); !ok {
		t.Errorf("Score value type=%T, want float64 (number)", firstRow["Score"])
	}
}

func createTestXlsx(t *testing.T, env []string) string {
	t.Helper()
	out := filepath.Join(t.TempDir(), "test.xlsx")
	spec := `{"sheets":[{"name":"Data","rows":[["Name","Score","Grade"],["Alice",95,"A"],["Bob",72,"C"],["Charlie",88,"B"],["Dave",60,"D"],["Eve",99,"A"]]}]}`
	r := run(t, env, "excel", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	return out
}

func TestExcel_Style(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "style", out, "--sheet", "Data", "--range", "A1:C1", "--bold", "--bg-color", "4472C4", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_InsertRows(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "insert-rows", out, "--sheet", "Data", "--after", "1", "--count", "2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_InsertCols(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "insert-cols", out, "--sheet", "Data", "--after", "1", "--count", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_DeleteRows(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "delete-rows", out, "--sheet", "Data", "--from", "2", "--count", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_DeleteCols(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "delete-cols", out, "--sheet", "Data", "--from", "3", "--count", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Sort(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "sort", out, "--sheet", "Data", "--range", "A1:C6", "--by-col", "2", "--ascending", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Freeze(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "freeze", out, "--sheet", "Data", "--cell", "A2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_MergeUnmerge(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	// Merge
	r := run(t, env, "excel", "merge", out, "--sheet", "Data", "--range", "A1:C1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("merge status=%v want ok", m["status"])
	}

	// Unmerge
	r2 := run(t, env, "excel", "unmerge", out, "--sheet", "Data", "--range", "A1:C1", "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	if m2["status"] != "ok" {
		t.Errorf("unmerge status=%v want ok", m2["status"])
	}
}

func TestExcel_SetColWidth(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "set-col-width", out, "--sheet", "Data", "--col", "1", "--width", "30", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_SetRowHeight(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "set-row-height", out, "--sheet", "Data", "--row", "1", "--height", "30", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Chart(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "chart", out, "--sheet", "Data", "--cell", "E1", "--type", "bar", "--data-range", "Data!A1:B6", "--title", "Scores", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_CondFormat(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "cond-format", out, "--sheet", "Data", "--range", "B2:B6", "--type", "cell", "--criteria", ">=", "--value", "90", "--bg-color", "FF0000", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_AutoFilter(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "auto-filter", out, "--sheet", "Data", "--range", "A1:C6", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_HideShowSheet(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	// Hide
	r := run(t, env, "excel", "hide-sheet", out, "--sheet", "Data", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("hide status=%v want ok", m["status"])
	}

	// Show
	r2 := run(t, env, "excel", "show-sheet", out, "--sheet", "Data", "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	if m2["status"] != "ok" {
		t.Errorf("show status=%v want ok", m2["status"])
	}
}

func TestExcel_AddImage(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	imgPath := createTestPNG(t)

	r := run(t, env, "excel", "add-image", out, "--sheet", "Data", "--cell", "E2", "--image", imgPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// --- Feature 2: formula ---

func TestExcel_Formula(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "formula", out, "--sheet", "Data", "--col", "B", "--func", "SUM", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "formulas")
	formulas, _ := m["formulas"].([]any)
	if len(formulas) != 1 {
		t.Fatalf("expected 1 formula result, got %d", len(formulas))
	}
	f0, _ := formulas[0].(map[string]any)
	if f0["formula"] != "=SUM(B2:B6)" {
		t.Errorf("formula=%v want =SUM(B2:B6)", f0["formula"])
	}
}

func TestExcel_Formula_Apply(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "formula", out, "--sheet", "Data", "--col", "B", "--func", "AVERAGE", "--apply", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["applied"] != true {
		t.Errorf("applied=%v want true", m["applied"])
	}

	// Read back to verify the formula was written.
	r2 := run(t, env, "excel", "read", out, "--sheet", "Data", "--typed", "--json")
	assertExit(t, r2, 0)
}

func TestExcel_Formula_Batch(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	spec := `[{"column":"B","func":"SUM"},{"column":"B","func":"AVERAGE"}]`
	r := run(t, env, "excel", "formula", out, "--sheet", "Data", "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	formulas, _ := m["formulas"].([]any)
	if len(formulas) != 2 {
		t.Fatalf("expected 2 formula results, got %d", len(formulas))
	}
}

func TestExcel_Formula_CountIf(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "formula", out, "--sheet", "Data", "--col", "B", "--func", "COUNTIF", "--criteria", ">=90", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	formulas, _ := m["formulas"].([]any)
	f0, _ := formulas[0].(map[string]any)
	formula, _ := f0["formula"].(string)
	if !strings.Contains(formula, "COUNTIF") {
		t.Errorf("formula=%v should contain COUNTIF", formula)
	}
}

// --- Feature 3: copy ---

func TestExcel_Copy(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)
	dst := filepath.Join(t.TempDir(), "copy.xlsx")

	r := run(t, env, "excel", "copy", out, "--output", dst, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify the copy can be read.
	if _, err := os.Stat(dst); err != nil {
		t.Errorf("copy file not created: %v", err)
	}
	r2 := run(t, env, "excel", "read", dst, "--sheet", "Data", "--json")
	assertExit(t, r2, 0)
}

func TestExcel_Copy_Force(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)
	dst := filepath.Join(t.TempDir(), "copy.xlsx")

	// First copy
	run(t, env, "excel", "copy", out, "--output", dst, "--json")
	// Second copy without --force should fail
	r := run(t, env, "excel", "copy", out, "--output", dst, "--json")
	assertExit(t, r, 7) // ENGINE_ERROR (file already exists)
	// With --force should succeed
	r2 := run(t, env, "excel", "copy", out, "--output", dst, "--force", "--json")
	assertExit(t, r2, 0)
}

// --- Feature 4: validation ---

func TestExcel_Validation_Dropdown(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "validation", out, "--sheet", "Data", "--range", "A2:A100",
		"--type", "list", "--list", "Alice,Bob,Carol", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["type"] != "list" {
		t.Errorf("type=%v want list", m["type"])
	}
}

func TestExcel_Validation_NumberRange(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "validation", out, "--sheet", "Data", "--range", "B2:B100",
		"--type", "whole", "--min", "0", "--max", "100", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Validation_Custom(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "validation", out, "--sheet", "Data", "--range", "A2:A100",
		"--type", "custom", "--formula", "=LEN(A2)>0", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Validation_WithMessages(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "validation", out, "--sheet", "Data", "--range", "B2:B100",
		"--type", "whole", "--min", "0", "--max", "100",
		"--error-msg", "Must be 0-100", "--error-title", "Invalid",
		"--prompt-msg", "Enter a number", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// --- Feature 5: style --spec batch ---

func TestExcel_Style_Batch(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	spec := `[{"range":"A1:C1","style":{"bold":true,"bgColor":"4472C4","textColor":"FFFFFF","align":"center"}},{"range":"A2:C6","style":{"numberFormat":"#,##0.00"}}]`
	r := run(t, env, "excel", "style", out, "--sheet", "Data", "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["entries"].(float64) != 2 {
		t.Errorf("entries=%v want 2", m["entries"])
	}
}

// --- Phase 2: column-formula, to-json, from-json, fill-range, copy-range, multi-chart, data-bar, icon-set ---

func TestExcel_ColumnFormula(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	spec := `{"column":"D","func":"IF","condition":"B{row}>80","valueIfTrue":"Pass","valueIfFalse":"Fail"}`
	r := run(t, env, "excel", "column-formula", out, "--sheet", "Data", "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "formula")
	requireField(t, m, "rowCount")
	if m["rowCount"].(float64) != 5 {
		t.Errorf("rowCount=%v want 5", m["rowCount"])
	}
}

func TestExcel_ColumnFormula_Apply(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	spec := `{"column":"D","func":"IF","condition":"B{row}>80","valueIfTrue":"Pass","valueIfFalse":"Fail"}`
	r := run(t, env, "excel", "column-formula", out, "--sheet", "Data", "--spec", spec, "--apply", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["applied"].(float64) != 5 {
		t.Errorf("applied=%v want 5", m["applied"])
	}
}

func TestExcel_ToJSON(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "to-json", out, "--sheet", "Data", "--with-headers", "--json")
	assertExit(t, r, 0)
	// to-json outputs a raw JSON array (not wrapped in {rows:...})
	var rows []map[string]any
	parseJSON(t, r.Stdout, &rows)
	if len(rows) != 5 {
		t.Fatalf("expected 5 data rows, got %d", len(rows))
	}
	if _, ok := rows[0]["Name"]; !ok {
		t.Errorf("first row missing 'Name' key: %v", rows[0])
	}
}

func TestExcel_FromJSON(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "fromjson.xlsx")

	data := `[{"Name":"Alice","Score":95},{"Name":"Bob","Score":80}]`
	r := run(t, env, "excel", "from-json", out, "--sheet", "Data", "--data", data, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["rows"].(float64) != 3 { // 1 header + 2 data
		t.Errorf("rows=%v want 3", m["rows"])
	}
	// Verify we can read it back
	r2 := run(t, env, "excel", "read", out, "--sheet", "Data", "--json")
	assertExit(t, r2, 0)
}

func TestExcel_FillRange(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "fill-range", out, "--sheet", "Data", "--range", "D1:D10", "--value", "N/A", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_CopyRange(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "copy-range", out, "--sheet", "Data", "--src", "A1:C3", "--dst", "D1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_MultiChart(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	spec := `{"type":"line","cell":"E2","title":"Chart","series":[{"name":"Score","catRange":"A2:A6","valRange":"B2:B6"},{"name":"Score Copy","catRange":"A2:A6","valRange":"B2:B6"}]}`
	r := run(t, env, "excel", "multi-chart", out, "--sheet", "Data", "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_DataBar(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "data-bar", out, "--sheet", "Data", "--range", "B2:B6", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_IconSet(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "icon-set", out, "--sheet", "Data", "--range", "B2:B6", "--style", "3Arrows", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_Style_Extended(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "style", out, "--sheet", "Data", "--range", "A1:C1", "--bold", "--strike", "--text-rotation", "45", "--indent", "2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestExcel_ColorScale(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "color-scale", out, "--sheet", "Data", "--range", "B2:B6", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["type"] != "2-color" {
		t.Errorf("type=%v want 2-color", m["type"])
	}
}

func TestExcel_ColorScale_3(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "color-scale", out, "--sheet", "Data", "--range", "B2:B6", "--mid-color", "FFEB84", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["type"] != "3-color" {
		t.Errorf("type=%v want 3-color", m["type"])
	}
}

func TestExcel_Hyperlink_SetGet(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	// Set hyperlink
	r := run(t, env, "excel", "hyperlink", out, "--sheet", "Data", "--cell", "A2", "--link", "https://example.com", "--display", "Click me", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Get hyperlink
	r2 := run(t, env, "excel", "hyperlink", out, "--sheet", "Data", "--cell", "A2", "--get", "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	if m2["exists"] != true {
		t.Errorf("exists=%v want true", m2["exists"])
	}
	if m2["link"] != "https://example.com" {
		t.Errorf("link=%v want https://example.com", m2["link"])
	}
}

func TestExcel_Hyperlink_Internal(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestXlsx(t, env)

	r := run(t, env, "excel", "hyperlink", out, "--sheet", "Data", "--cell", "A3", "--link", "Data!A1", "--type", "Location", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// ============================================================================
// WORD NEW COMMANDS
// ============================================================================

func TestWord_Create(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "new.docx")

	r := run(t, env, "word", "create", out, "--title", "Test Doc", "--author", "Test", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if _, err := os.Stat(out); err != nil {
		t.Errorf("file not created: %v", err)
	}
}

func TestWord_Create_AddParagraph_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "para.docx")

	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "add-paragraph", out, "--text", "Hello World", "--style", "Normal", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Read back
	r2 := run(t, env, "word", "read", out, "--format", "text", "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	text, _ := read["text"].(string)
	if !containsStr(text, "Hello World") {
		t.Errorf("read back text=%q does not contain 'Hello World'", text)
	}
}

func TestWord_Create_AddHeading_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "heading.docx")

	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "add-heading", out, "--text", "Chapter 1", "--level", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Read back headings
	r2 := run(t, env, "word", "headings", out, "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	headings, _ := read["headings"].([]any)
	if len(headings) == 0 {
		t.Error("expected at least 1 heading")
	}
}

func TestWord_Create_AddTable_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "table.docx")

	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "add-table", out, "--rows", `[["Name","Age"],["Alice","30"],["Bob","25"]]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify the file is valid by reading it back
	r2 := run(t, env, "word", "read", out, "--format", "paragraphs", "--json")
	assertExit(t, r2, 0)
}

func TestWord_Create_AddPageBreak(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "pagebreak.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Page 1", "--json")

	r := run(t, env, "word", "add-page-break", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestWord_Create_AddImage(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "img.docx")

	run(t, env, "word", "create", out, "--json")

	imgPath := createTestPNG(t)

	r := run(t, env, "word", "add-image", out, "--image", imgPath, "--width", "100", "--height", "100", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// --- Read with tables ---

func TestWord_Read_WithTables_JSON(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--with-tables", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "elements")
	paras, _ := m["elements"].([]any)
	if len(paras) == 0 {
		t.Fatal("expected body elements")
	}
	// Find a table entry
	found := false
	for _, p := range paras {
		elem, ok := p.(map[string]any)
		if ok && elem["type"] == "table" {
			found = true
			rows, _ := elem["rows"].([]any)
			if len(rows) == 0 {
				t.Error("table should have rows")
			}
			break
		}
	}
	if !found {
		t.Error("expected at least one table element in output")
	}
}

func TestWord_Read_WithTables_Markdown(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--with-tables", "--format", "markdown", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "markdown")
	md, _ := m["markdown"].(string)
	if !strings.Contains(md, "|") {
		t.Error("markdown output should contain table pipe characters")
	}
}

func TestWord_Read_WithTables_Text(t *testing.T) {
	env := setupHome(t, "read-only")
	r := run(t, env, "word", "read", fixture("sample.docx"), "--with-tables", "--format", "text", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "text")
}

// --- Delete ---

func TestWord_Delete(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "del.docx")

	// Create doc with content
	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Keep", "--style", "Normal", "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Remove Me", "--style", "Normal", "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Also Keep", "--style", "Normal", "--json")

	// Delete element at index 1 ("Remove Me")
	r := run(t, env, "word", "delete", out, "--index", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify
	r2 := run(t, env, "word", "read", out, "--format", "text", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	text, _ := read["text"].(string)
	if strings.Contains(text, "Remove Me") {
		t.Error("deleted text should not appear in output")
	}
	if !containsStr(text, "Keep") {
		t.Error("remaining text should still be present")
	}
}

func TestWord_Delete_Table(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "del-tbl.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-table", out, "--rows", `[["A","B"],["1","2"]]`, "--json")

	// Read with tables to find the table's index
	r := run(t, env, "word", "read", out, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r.Stdout, &read)
	elems, _ := read["elements"].([]any)
	tableIdx := -1
	for i, p := range elems {
		elem := p.(map[string]any)
		if elem["type"] == "table" {
			tableIdx = i
			break
		}
	}
	if tableIdx < 0 {
		t.Fatal("table not found in read output")
	}

	// Delete the table
	r2 := run(t, env, "word", "delete", out, "--index", fmt.Sprintf("%d", tableIdx), "--json")
	assertExit(t, r2, 0)
}

// --- Insert before / after ---

func TestWord_InsertBefore(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "ins-before.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Existing", "--json")

	r := run(t, env, "word", "insert-before", out, "--index", "0", "--type", "paragraph", "--text", "Before Existing", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify ordering
	r2 := run(t, env, "word", "read", out, "--with-tables", "--format", "paragraphs", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elems, _ := read["elements"].([]any)
	if len(elems) < 2 {
		t.Fatal("expected at least 2 elements")
	}
	// First should be "Before Existing"
	first := elems[0].(map[string]any)
	if first["text"] != "Before Existing" {
		t.Errorf("first element text=%v want 'Before Existing'", first["text"])
	}
}

func TestWord_InsertAfter(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "ins-after.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Existing", "--json")

	r := run(t, env, "word", "insert-after", out, "--index", "0", "--type", "paragraph", "--text", "After Existing", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify ordering
	r2 := run(t, env, "word", "read", out, "--with-tables", "--format", "paragraphs", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elems, _ := read["elements"].([]any)
	if len(elems) < 2 {
		t.Fatal("expected at least 2 elements")
	}
	// Second should be "After Existing"
	second := elems[1].(map[string]any)
	if second["text"] != "After Existing" {
		t.Errorf("second element text=%v want 'After Existing'", second["text"])
	}
}

func TestWord_InsertBefore_Heading(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "ins-heading.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Body", "--json")

	r := run(t, env, "word", "insert-before", out, "--index", "0", "--type", "heading", "--text", "New Chapter", "--level", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestWord_InsertBefore_Table(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "ins-table.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Body", "--json")

	rows := `[["Name","Age"],["Alice","30"]]`
	r := run(t, env, "word", "insert-before", out, "--index", "0", "--type", "table", "--rows", rows, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// --- Update table cell ---

func TestWord_UpdateTableCell(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "utc.docx")

	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-table", out, "--rows", `[["Name","Score"],["Alice","95"],["Bob","82"]]`, "--json")

	// Read with tables to find table index
	r := run(t, env, "word", "read", out, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r.Stdout, &read)
	elems, _ := read["elements"].([]any)
	tableIdx := -1
	for i, p := range elems {
		elem := p.(map[string]any)
		if elem["type"] == "table" {
			tableIdx = i
			break
		}
	}
	if tableIdx < 0 {
		t.Fatal("table not found")
	}

	// Update Alice's score from 95 to 100
	r2 := run(t, env, "word", "update-table-cell", out,
		"--table-index", fmt.Sprintf("%d", tableIdx),
		"--row", "1", "--col", "1", "--value", "100", "--json")
	assertExit(t, r2, 0)
	var m map[string]any
	parseJSON(t, r2.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify
	r3 := run(t, env, "word", "read", out, "--with-tables", "--json")
	var verify map[string]any
	parseJSON(t, r3.Stdout, &verify)
	paras2, _ := verify["elements"].([]any)
	for _, p := range paras2 {
		elem := p.(map[string]any)
		if elem["type"] == "table" {
			rows, _ := elem["rows"].([]any)
			if len(rows) >= 2 {
				row1, _ := rows[1].([]any)
				if len(row1) >= 2 && row1[1] != "100" {
					t.Errorf("cell value=%v want 100", row1[1])
				}
			}
			break
		}
	}
}

// --- Merge ---

func TestWord_Merge(t *testing.T) {
	env := setupHome(t, "write")

	// Create two separate documents
	doc1 := filepath.Join(t.TempDir(), "merge1.docx")
	doc2 := filepath.Join(t.TempDir(), "merge2.docx")
	out := filepath.Join(t.TempDir(), "merged.docx")

	run(t, env, "word", "create", doc1, "--json")
	run(t, env, "word", "add-paragraph", doc1, "--text", "Document One", "--json")
	run(t, env, "word", "add-heading", doc1, "--text", "Section A", "--level", "1", "--json")

	run(t, env, "word", "create", doc2, "--json")
	run(t, env, "word", "add-paragraph", doc2, "--text", "Document Two", "--json")
	run(t, env, "word", "add-table", doc2, "--rows", `[["X","Y"],["1","2"]]`, "--json")

	r := run(t, env, "word", "merge", doc1, doc2, "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify merged content
	r2 := run(t, env, "word", "read", out, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	paras, _ := read["elements"].([]any)
	// Should have: Document One + Section A + Document Two + Table = 4 elements
	if len(paras) < 3 {
		t.Errorf("expected >= 3 elements in merged doc, got %d", len(paras))
	}
}

func TestWord_Merge_ThreeFiles(t *testing.T) {
	env := setupHome(t, "write")

	docs := make([]string, 3)
	for i := range docs {
		docs[i] = filepath.Join(t.TempDir(), fmt.Sprintf("merge3_%d.docx", i))
		run(t, env, "word", "create", docs[i], "--json")
		run(t, env, "word", "add-paragraph", docs[i], "--text", fmt.Sprintf("Part %d", i+1), "--json")
	}
	out := filepath.Join(t.TempDir(), "merged3.docx")

	r := run(t, env, "word", "merge", docs[0], docs[1], docs[2], "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

// ---------------------------------------------------------------------------
// word search
// ---------------------------------------------------------------------------

func TestWord_Search(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "search.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-paragraph", doc, "--text", "Revenue increased by 20%", "--json")
	run(t, env, "word", "add-paragraph", doc, "--text", "Expenses remained flat", "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["Item","Revenue"],["Product A","$1M"],["Product B","$2M"]]`, "--json")

	r := run(t, env, "word", "search", doc, "--keyword", "Revenue", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["keyword"] != "Revenue" {
		t.Errorf("keyword=%v want Revenue", m["keyword"])
	}
	hits := int(m["hits"].(float64))
	if hits < 2 {
		t.Errorf("expected >= 2 hits for 'Revenue', got %d", hits)
	}
}

func TestWord_Search_NoMatch(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "search-nomatch.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-paragraph", doc, "--text", "Hello World", "--json")

	r := run(t, env, "word", "search", doc, "--keyword", "ZZZZNOTFOUND", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if int(m["hits"].(float64)) != 0 {
		t.Errorf("hits=%v want 0", m["hits"])
	}
}

func TestWord_Search_MissingKeyword(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "search-err.docx")
	run(t, env, "word", "create", doc, "--json")

	r := run(t, env, "word", "search", doc, "--json")
	assertExit(t, r, 2) // VALIDATION_ERROR
}

// ---------------------------------------------------------------------------
// word add-table-rows
// ---------------------------------------------------------------------------

func TestWord_AddTableRows(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "addrows.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B"],["1","2"]]`, "--json")

	r := run(t, env, "word", "add-table-rows", doc, "--table-index", "0", "--rows", `[["3","4"],["5","6"]]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back
	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	if len(rows) != 4 {
		t.Errorf("expected 4 rows after add-table-rows, got %d", len(rows))
	}
}

func TestWord_AddTableRows_AtPosition(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "addrows-pos.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B"],["1","2"],["3","4"]]`, "--json")

	r := run(t, env, "word", "add-table-rows", doc, "--table-index", "0", "--position", "1", "--rows", `[["X","Y"]]`, "--json")
	assertExit(t, r, 0)

	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	if len(rows) != 4 {
		t.Errorf("expected 4 rows, got %d", len(rows))
	}
}

// ---------------------------------------------------------------------------
// word delete-table-rows
// ---------------------------------------------------------------------------

func TestWord_DeleteTableRows(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "delrows.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B"],["1","2"],["3","4"],["5","6"]]`, "--json")

	r := run(t, env, "word", "delete-table-rows", doc, "--table-index", "0", "--start-row", "2", "--end-row", "3", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back — should have 2 rows left (header + row 1)
	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	if len(rows) != 2 {
		t.Errorf("expected 2 rows after delete, got %d", len(rows))
	}
}

// ---------------------------------------------------------------------------
// word add-table-cols
// ---------------------------------------------------------------------------

func TestWord_AddTableCols(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "addcols.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B"],["1","2"],["3","4"]]`, "--json")

	r := run(t, env, "word", "add-table-cols", doc, "--table-index", "0", "--values", `["X","Y","Z"]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back
	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	// Each row should now have 3 cells
	firstRow := rows[0].([]any)
	if len(firstRow) != 3 {
		t.Errorf("expected 3 cols, got %d", len(firstRow))
	}
}

// ---------------------------------------------------------------------------
// word delete-table-cols
// ---------------------------------------------------------------------------

func TestWord_DeleteTableCols(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "delcols.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B","C"],["1","2","3"]]`, "--json")

	r := run(t, env, "word", "delete-table-cols", doc, "--table-index", "0", "--start-col", "1", "--end-col", "2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back — should have 1 column left
	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	firstRow := rows[0].([]any)
	if len(firstRow) != 1 {
		t.Errorf("expected 1 col, got %d", len(firstRow))
	}
}

// ---------------------------------------------------------------------------
// word update-table (batch cell update)
// ---------------------------------------------------------------------------

func TestWord_UpdateTable_Batch(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "batchupdate.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B","C"],["1","2","3"],["4","5","6"]]`, "--json")

	spec := `[{"row":0,"col":0,"value":"NEW_A"},{"row":1,"col":2,"value":"NEW_3"},{"row":2,"col":1,"value":"NEW_5"}]`
	r := run(t, env, "word", "update-table", doc, "--table-index", "0", "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if int(m["updated"].(float64)) != 3 {
		t.Errorf("updated=%v want 3", m["updated"])
	}

	// Verify: read back
	r2 := run(t, env, "word", "read", doc, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	elements, _ := read["elements"].([]any)
	tbl := elements[0].(map[string]any)
	rows, _ := tbl["rows"].([]any)
	firstRow := rows[0].([]any)
	if firstRow[0].(string) != "NEW_A" {
		t.Errorf("cell[0][0]=%v want NEW_A", firstRow[0])
	}
}

func TestWord_UpdateTable_MissingSpec(t *testing.T) {
	env := setupHome(t, "write")
	doc := filepath.Join(t.TempDir(), "batchupdate-err.docx")
	run(t, env, "word", "create", doc, "--json")
	run(t, env, "word", "add-table", doc, "--rows", `[["A","B"]]`, "--json")

	r := run(t, env, "word", "update-table", doc, "--table-index", "0", "--json")
	assertExit(t, r, 2) // VALIDATION_ERROR
}

// ============================================================================
// PPT NEW COMMANDS
// ============================================================================

func TestPPT_Create(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "new.pptx")

	r := run(t, env, "ppt", "create", out, "--title", "My Talk", "--author", "Test", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if _, err := os.Stat(out); err != nil {
		t.Errorf("file not created: %v", err)
	}
}

func TestPPT_Create_CountSlides(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "count.pptx")

	run(t, env, "ppt", "create", out, "--json")

	r := run(t, env, "ppt", "count", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["slides"].(float64) != 1 {
		t.Errorf("slides=%v want 1", m["slides"])
	}
}

func TestPPT_Create_AddSlide_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "addslide.pptx")

	run(t, env, "ppt", "create", out, "--json")

	r := run(t, env, "ppt", "add-slide", out, "--title", "Slide 2", "--bullets", `["Point A","Point B"]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Count slides
	r2 := run(t, env, "ppt", "count", out, "--json")
	var count map[string]any
	parseJSON(t, r2.Stdout, &count)
	if count["slides"].(float64) != 2 {
		t.Errorf("slides=%v want 2", count["slides"])
	}
}

func TestPPT_Create_SetContent_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "setcontent.pptx")

	run(t, env, "ppt", "create", out, "--title", "Original", "--json")

	r := run(t, env, "ppt", "set-content", out, "--slide", "1", "--title", "Updated", "--bullets", `["New bullet"]`, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Read back
	r2 := run(t, env, "ppt", "read", out, "--format", "outline", "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	slides, _ := read["slides"].([]any)
	if len(slides) == 0 {
		t.Fatal("no slides found")
	}
	slide := slides[0].(map[string]any)
	if slide["title"] != "Updated" {
		t.Errorf("title=%v want Updated", slide["title"])
	}
}

func TestPPT_Create_SetNotes_ReadBack(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "notes.pptx")

	run(t, env, "ppt", "create", out, "--json")

	r := run(t, env, "ppt", "set-notes", out, "--slide", "1", "--notes", "Remember to smile!", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Read back with notes
	r2 := run(t, env, "ppt", "read", out, "--format", "outline", "--with-notes", "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	slides, _ := read["slides"].([]any)
	if len(slides) == 0 {
		t.Fatal("no slides found")
	}
	slide := slides[0].(map[string]any)
	if slide["notes"] != "Remember to smile!" {
		t.Errorf("notes=%v want 'Remember to smile!'", slide["notes"])
	}
}

func TestPPT_Create_DeleteSlide(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "delete.pptx")

	run(t, env, "ppt", "create", out, "--json")
	run(t, env, "ppt", "add-slide", out, "--title", "Slide 2", "--json")

	// Delete slide 1
	r := run(t, env, "ppt", "delete-slide", out, "--slide", "1", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Count should be 1
	r2 := run(t, env, "ppt", "count", out, "--json")
	var count map[string]any
	parseJSON(t, r2.Stdout, &count)
	if count["slides"].(float64) != 1 {
		t.Errorf("slides=%v want 1", count["slides"])
	}
}

func TestPPT_Create_Reorder(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "reorder.pptx")

	run(t, env, "ppt", "create", out, "--json")
	run(t, env, "ppt", "add-slide", out, "--title", "Second", "--json")
	run(t, env, "ppt", "add-slide", out, "--title", "Third", "--json")

	// Reorder: 3,1,2
	r := run(t, env, "ppt", "reorder", out, "--order", "3,1,2", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestPPT_Create_AddImage(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "img.pptx")

	run(t, env, "ppt", "create", out, "--json")

	imgPath := createTestPNG(t)

	r := run(t, env, "ppt", "add-image", out, "--slide", "1", "--image", imgPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestPPT_Build(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "built.pptx")
	spec := `{"title":"Q2 Review","author":"Agent","slides":[{"title":"Overview","bullets":["Revenue up 20%","Users doubled"]},{"title":"Roadmap","bullets":["Q3: launch","Q4: scale"]},{"title":"Q&A","notes":"Remember to thank the team"}]}`

	// 1. Build a deck from spec
	r := run(t, env, "ppt", "build", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if int(m["slides"].(float64)) != 3 {
		t.Errorf("slides=%v want 3", m["slides"])
	}

	// 2. Read back and verify content
	r2 := run(t, env, "ppt", "read", out, "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	slides, ok := m2["slides"].([]any)
	if !ok {
		t.Fatalf("slides not an array: %T", m2["slides"])
	}
	if len(slides) != 3 {
		t.Fatalf("expected 3 slides, got %d", len(slides))
	}

	// Check slide titles
	first, _ := slides[0].(map[string]any)
	if first["title"] != "Overview" {
		t.Errorf("slide 1 title=%v want Overview", first["title"])
	}
	second, _ := slides[1].(map[string]any)
	if second["title"] != "Roadmap" {
		t.Errorf("slide 2 title=%v want Roadmap", second["title"])
	}
	third, _ := slides[2].(map[string]any)
	if third["title"] != "Q&A" {
		t.Errorf("slide 3 title=%v want Q&A", third["title"])
	}

	// 3. Verify notes for slide 3
	r3 := run(t, env, "ppt", "read", out, "--with-notes", "--json")
	assertExit(t, r3, 0)
	var m3 map[string]any
	parseJSON(t, r3.Stdout, &m3)
	slides3, _ := m3["slides"].([]any)
	thirdWithNotes, _ := slides3[2].(map[string]any)
	if thirdWithNotes["notes"] != "Remember to thank the team" {
		t.Errorf("slide 3 notes=%v want 'Remember to thank the team'", thirdWithNotes["notes"])
	}
}

// ============================================================================
// PDF NEW COMMANDS
// ============================================================================

func TestPDF_Bookmarks(t *testing.T) {
	env := setupHome(t, "read-only")

	r := run(t, env, "pdf", "bookmarks", fixture("sample.pdf"), "--json")
	// This may return empty bookmarks if the fixture has none - that's fine
	assertExit(t, r, 0)
}

func createTestPDF(t *testing.T, env []string) string {
	t.Helper()
	out := filepath.Join(t.TempDir(), "multi.pdf")
	// Use pdfcpu to create a simple multi-page PDF via the split/merge workflow
	// Instead, we'll use the fixture and make a copy
	src := fixture("sample.pdf")
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(out, data, 0644); err != nil {
		t.Fatal(err)
	}
	return out
}

func TestPDF_InsertBlank(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestPDF(t, env)
	outPath := filepath.Join(t.TempDir(), "blank.pdf")

	r := run(t, env, "pdf", "insert-blank", out, "--after", "1", "--count", "1", "--output", outPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify: output file should now have 2 pages (original 1 + 1 blank)
	r2 := run(t, env, "pdf", "pages", outPath, "--json")
	var pages map[string]any
	parseJSON(t, r2.Stdout, &pages)
	if pages["pages"].(float64) != 2 {
		t.Errorf("expected 2 pages after insert-blank, got %v", pages["pages"])
	}
}

func TestPDF_SetMeta(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestPDF(t, env)
	outPath := filepath.Join(t.TempDir(), "meta.pdf")

	r := run(t, env, "pdf", "set-meta", out, "--title", "New Title", "--author", "Test Author", "--output", outPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify metadata via info
	r2 := run(t, env, "pdf", "info", outPath, "--json")
	var info map[string]any
	parseJSON(t, r2.Stdout, &info)
	if info["title"] != "New Title" {
		t.Errorf("title=%v want 'New Title'", info["title"])
	}
}

func TestPDF_Reorder(t *testing.T) {
	env := setupHome(t, "write")
	// Use sample-2.pdf which has 2 pages
	src := fixture("sample-2.pdf")
	out := filepath.Join(t.TempDir(), "reorder-src.pdf")
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(out, data, 0644); err != nil {
		t.Fatal(err)
	}
	outPath := filepath.Join(t.TempDir(), "reorder.pdf")

	r := run(t, env, "pdf", "reorder", out, "--order", "2,1", "--output", outPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestPDF_AddBookmarks(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestPDF(t, env)
	outPath := filepath.Join(t.TempDir(), "bookmarked.pdf")

	r := run(t, env, "pdf", "add-bookmarks", out, "--spec", `[{"title":"Introduction","page":1}]`, "--output", outPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify bookmarks can be read back
	r2 := run(t, env, "pdf", "bookmarks", outPath, "--json")
	assertExit(t, r2, 0)
	var bm map[string]any
	parseJSON(t, r2.Stdout, &bm)
	bookmarks, _ := bm["bookmarks"].([]any)
	if len(bookmarks) < 1 {
		t.Error("expected at least 1 bookmark after add-bookmarks")
	}
}

func TestPDF_StampImage(t *testing.T) {
	env := setupHome(t, "write")
	out := createTestPDF(t, env)
	outPath := filepath.Join(t.TempDir(), "stamped.pdf")
	imgPath := createTestPNG(t)

	r := run(t, env, "pdf", "stamp-image", out, "--image", imgPath, "--pages", "1", "--output", outPath, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify output file exists and is valid PDF
	r2 := run(t, env, "pdf", "pages", outPath, "--json")
	assertExit(t, r2, 0)
}

// ============================================================================
// HELPERS
// ============================================================================

// createTestPNG creates a minimal valid 1x1 PNG file and returns its path.
func createTestPNG(t *testing.T) string {
	t.Helper()
	imgPath := filepath.Join(t.TempDir(), "test.png")
	f, err := os.Create(imgPath)
	if err != nil {
		t.Fatal(err)
	}
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{R: 255, A: 255})
	if err := png.Encode(f, img); err != nil {
		_ = f.Close()
		t.Fatal(err)
	}
	_ = f.Close()
	return imgPath
}

// containsStr is a simple substring check.
func containsStr(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(s) > 0 && containsSubstr(s, substr))
}

func containsSubstr(s, sub string) bool {
	for i := 0; i <= len(s)-len(sub); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}

// ============================================================================
// PDF NEW COMMANDS: search, create, replace
// ============================================================================

func TestPDF_Search(t *testing.T) {
	env := setupHome(t, "read-only")

	// Use the sample PDF fixture
	r := run(t, env, "pdf", "search", fixture("sample.pdf"), "Lorem", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "count")
	requireField(t, m, "matches")
	if m["keyword"] != "Lorem" {
		t.Errorf("keyword=%v want Lorem", m["keyword"])
	}
}

func TestPDF_Search_NotFound(t *testing.T) {
	env := setupHome(t, "read-only")

	r := run(t, env, "pdf", "search", fixture("sample.pdf"), "zzznonexistent", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if int(m["count"].(float64)) != 0 {
		t.Errorf("count=%v want 0", m["count"])
	}
}

func TestPDF_Search_WithFlags(t *testing.T) {
	env := setupHome(t, "read-only")

	r := run(t, env, "pdf", "search", fixture("sample.pdf"), "Lorem", "--case-sensitive", "--context", "10", "--limit", "5", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	requireField(t, m, "matches")
}

func TestPDF_Create(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "created.pdf")

	spec := `{"pages":[{"text":"Hello World"},{"text":"Page Two Content"}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["pages"].(float64) != 2 {
		t.Errorf("pages=%v want 2", m["pages"])
	}

	// Verify: read back the created PDF
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	pages, _ := read["pages"].([]any)
	if len(pages) != 2 {
		t.Fatalf("expected 2 pages in readback, got %d", len(pages))
	}
}

func TestPDF_Create_WithMetadata(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "meta-create.pdf")

	spec := `{"pages":[{"text":"Content"}],"title":"My Doc","author":"Agent"}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify metadata via info
	r2 := run(t, env, "pdf", "info", out, "--json")
	assertExit(t, r2, 0)
	var info map[string]any
	parseJSON(t, r2.Stdout, &info)
	if info["title"] != "My Doc" {
		t.Errorf("title=%v want 'My Doc'", info["title"])
	}
	if info["author"] != "Agent" {
		t.Errorf("author=%v want 'Agent'", info["author"])
	}
}

func TestPDF_Create_MissingSpec(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "no-spec.pdf")

	r := run(t, env, "pdf", "create", out, "--json")
	assertExit(t, r, 2) // VALIDATION_ERROR
}

func TestPDF_Replace(t *testing.T) {
	env := setupHome(t, "write")

	// First create a PDF with known text
	src := filepath.Join(t.TempDir(), "replace-src.pdf")
	spec := `{"pages":[{"text":"Hello World Goodbye"}]}`
	run(t, env, "pdf", "create", src, "--spec", spec, "--json")

	out := filepath.Join(t.TempDir(), "replaced.pdf")

	r := run(t, env, "pdf", "replace", src, "--find", "Hello", "--replace", "Goodbye", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--text-only", "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	text, _ := read["text"].(string)
	if !strings.Contains(text, "Goodbye") {
		t.Errorf("replaced text missing 'Goodbye': %q", text)
	}
}

func TestPDF_Replace_BatchPairs(t *testing.T) {
	env := setupHome(t, "write")

	// Create PDF with known text
	src := filepath.Join(t.TempDir(), "batch-src.pdf")
	spec := `{"pages":[{"text":"Alice and Bob"}]}`
	run(t, env, "pdf", "create", src, "--spec", spec, "--json")

	out := filepath.Join(t.TempDir(), "batch-replaced.pdf")

	pairs := `[{"find":"Alice","replace":"Charlie"},{"find":"Bob","replace":"Dave"}]`
	r := run(t, env, "pdf", "replace", src, "--pairs", pairs, "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestPDF_Replace_DefaultOutput(t *testing.T) {
	env := setupHome(t, "write")

	// Create PDF with known text
	src := filepath.Join(t.TempDir(), "default-out.pdf")
	spec := `{"pages":[{"text":"Replace Me"}]}`
	run(t, env, "pdf", "create", src, "--spec", spec, "--json")

	// Don't specify --output, should default to <input>.replaced.pdf
	r := run(t, env, "pdf", "replace", src, "--find", "Replace", "--replace", "Done", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	// Verify output path in JSON response ends with .replaced.pdf
	outPath, _ := m["output"].(string)
	if !strings.HasSuffix(outPath, ".replaced.pdf") {
		t.Errorf("output=%q should end with .replaced.pdf", outPath)
	}
	// Verify the output file exists
	if _, err := os.Stat(outPath); err != nil {
		t.Errorf("default output file not found at %s: %v", outPath, err)
	}
}

// ---------------------------------------------------------------------------
// PDF AddText E2E
// ---------------------------------------------------------------------------

func TestPDF_AddText(t *testing.T) {
	env := setupHome(t, "write")

	// Create a source PDF
	src := filepath.Join(t.TempDir(), "addtext-src.pdf")
	spec := `{"pages":[{"text":"Original Content"}]}`
	run(t, env, "pdf", "create", src, "--spec", spec, "--json")

	out := filepath.Join(t.TempDir(), "addtext-out.pdf")
	r := run(t, env, "pdf", "add-text", src, "--text", "DRAFT", "--x", "200", "--y", "400", "--font-size", "24", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["overlays"].(float64) != 1 {
		t.Errorf("overlays=%v want 1", m["overlays"])
	}

	// Verify: check that the output file exists and has content
	info, err := os.Stat(out)
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("output file is empty")
	}
}

func TestPDF_AddText_WithColor(t *testing.T) {
	env := setupHome(t, "write")

	src := filepath.Join(t.TempDir(), "color-src.pdf")
	run(t, env, "pdf", "create", src, "--spec", `{"pages":[{"text":"Page"}]}`, "--json")

	out := filepath.Join(t.TempDir(), "color-out.pdf")
	r := run(t, env, "pdf", "add-text", src, "--text", "RED", "--x", "100", "--y", "300", "--color", "1 0 0", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestPDF_AddText_MissingOutput(t *testing.T) {
	env := setupHome(t, "write")

	src := filepath.Join(t.TempDir(), "no-output-src.pdf")
	run(t, env, "pdf", "create", src, "--spec", `{"pages":[{"text":"Page"}]}`, "--json")

	r := run(t, env, "pdf", "add-text", src, "--text", "TEST")
	assertExit(t, r, 2) // missing --output → validation error
}

func TestPDF_Create_WithFontSize(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "fontsize.pdf")

	spec := `{"pages":[{"text":"Large Text","fontSize":24}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify: read back the created PDF
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
	var read map[string]any
	parseJSON(t, r2.Stdout, &read)
	pages, _ := read["pages"].([]any)
	if len(pages) != 1 {
		t.Fatalf("expected 1 page in readback, got %d", len(pages))
	}
}

func TestPDF_Create_BoldAndColor(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "boldcolor.pdf")

	spec := `{"pages":[{"text":"Bold Red","bold":true,"fontSize":18,"color":"1 0 0"},{"text":"Normal Blue","color":"0 0 1"}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 2 {
		t.Errorf("pages=%v want 2", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

// ---------------------------------------------------------------------------
// PDF Typography E2E Tests
// ---------------------------------------------------------------------------

func TestPDF_Create_TimesSerif(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "times-serif.pdf")

	spec := `{"pages":[{"text":"Serif text","font":"Times"},{"text":"Bold serif","font":"Times","bold":true}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 2 {
		t.Errorf("pages=%v want 2", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_Create_CourierMono(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "courier-mono.pdf")

	spec := `{"pages":[{"text":"Monospace text","font":"Courier"},{"text":"Bold mono","font":"Courier","bold":true,"italic":true}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 2 {
		t.Errorf("pages=%v want 2", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_Create_Alignment(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "aligned.pdf")

	spec := `{"pages":[{"text":"Left aligned","align":"left"},{"text":"Center aligned","align":"center"},{"text":"Right aligned","align":"right"}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 3 {
		t.Errorf("pages=%v want 3", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_Create_Underline(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "underline.pdf")

	spec := `{"pages":[{"text":"Underlined text","underline":true}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 1 {
		t.Errorf("pages=%v want 1", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_Create_GlobalDefaults(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "global-defaults.pdf")

	spec := `{"font":"Times","fontSize":14,"align":"center","lineHeight":1.8,"margin":50,"pages":[{"text":"Page 1 with global defaults"},{"text":"Page 2 with global defaults"}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 2 {
		t.Errorf("pages=%v want 2", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_Create_MixedFonts(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "mixed-fonts.pdf")

	spec := `{"pages":[{"text":"Sans-serif","font":"Helvetica"},{"text":"Serif","font":"Times"},{"text":"Monospace","font":"Courier"},{"text":"Bold Serif","font":"Times","bold":true},{"text":"Italic Mono","font":"Courier","italic":true}]}`
	r := run(t, env, "pdf", "create", out, "--spec", spec, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["pages"].(float64) != 5 {
		t.Errorf("pages=%v want 5", m["pages"])
	}

	// Verify: read back
	r2 := run(t, env, "pdf", "read", out, "--json")
	assertExit(t, r2, 0)
}

func TestPDF_AddText_FontFlags(t *testing.T) {
	env := setupHome(t, "write")

	// Create a source PDF
	src := filepath.Join(t.TempDir(), "font-src.pdf")
	run(t, env, "pdf", "create", src, "--spec", `{"pages":[{"text":"Original Content"}]}`, "--json")

	out := filepath.Join(t.TempDir(), "font-out.pdf")
	r := run(t, env, "pdf", "add-text", src, "--text", "Serif Overlay", "--font", "Times", "--bold", "--italic", "--underline", "--x", "200", "--y", "400", "--font-size", "24", "--output", out, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if m["overlays"].(float64) != 1 {
		t.Errorf("overlays=%v want 1", m["overlays"])
	}
}

// ---------------------------------------------------------------------------
// Word style commands
// ---------------------------------------------------------------------------

func TestWord_Style_Paragraph(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "styled-para.docx")

	// Create a document with a paragraph
	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-paragraph", out, "--text", "Important Note", "--style", "Heading1", "--json")

	// Style the paragraph (element 0 is the heading)
	r := run(t, env, "word", "style", out,
		"--index", "0",
		"--bold", "--color", "FF0000", "--font-size", "16",
		"--align", "center",
		"--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestWord_Style_Paragraph_MissingIndex(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "err.docx")
	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "style", out, "--bold")
	assertExit(t, r, 2) // missing --index
}

func TestWord_Style_Paragraph_NoStyleFlags(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "err.docx")
	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "style", out, "--index", "0")
	assertExit(t, r, 2) // no style flags
}

func TestWord_Style_TableCells(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "styled-table.docx")

	// Create doc with table
	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-table", out, "--rows", `[["Name","Status"],["Alice","PASS"],["Bob","FAIL"]]`, "--json")

	// Find table index
	r := run(t, env, "word", "read", out, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r.Stdout, &read)
	elems, _ := read["elements"].([]any)
	tableIdx := -1
	for i, p := range elems {
		elem := p.(map[string]any)
		if elem["type"] == "table" {
			tableIdx = i
			break
		}
	}
	if tableIdx < 0 {
		t.Fatal("table not found")
	}

	// Style the header row (row 0)
	r2 := run(t, env, "word", "style-table", out,
		"--table-index", fmt.Sprintf("%d", tableIdx),
		"--start-row", "0", "--end-row", "0",
		"--bg-color", "003366", "--bold", "--color", "FFFFFF",
		"--json")
	assertExit(t, r2, 0)
	var m map[string]any
	parseJSON(t, r2.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
}

func TestWord_Style_TableBatch(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "batch-styled.docx")

	// Create doc with table
	run(t, env, "word", "create", out, "--json")
	run(t, env, "word", "add-table", out, "--rows", `[["Name","Status"],["Alice","PASS"],["Bob","FAIL"]]`, "--json")

	// Find table index
	r := run(t, env, "word", "read", out, "--with-tables", "--json")
	var read map[string]any
	parseJSON(t, r.Stdout, &read)
	elems, _ := read["elements"].([]any)
	tableIdx := -1
	for i, p := range elems {
		elem := p.(map[string]any)
		if elem["type"] == "table" {
			tableIdx = i
			break
		}
	}
	if tableIdx < 0 {
		t.Fatal("table not found")
	}

	// Batch style: header row dark blue, row 2 red bg
	spec := fmt.Sprintf(`[{"startRow":0,"endRow":0,"bgColor":"003366","bold":true,"color":"FFFFFF"},{"startRow":2,"bgColor":"FF0000","color":"FFFFFF"}]`)
	r2 := run(t, env, "word", "style-table", out,
		"--table-index", fmt.Sprintf("%d", tableIdx),
		"--spec", spec,
		"--json")
	assertExit(t, r2, 0)
	var m map[string]any
	parseJSON(t, r2.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}
	if styled, ok := m["styled"].(float64); !ok || styled != 2 {
		t.Errorf("styled=%v want 2", m["styled"])
	}
}

func TestWord_Style_TableMissingIndex(t *testing.T) {
	env := setupHome(t, "write")
	out := filepath.Join(t.TempDir(), "err.docx")
	run(t, env, "word", "create", out, "--json")

	r := run(t, env, "word", "style-table", out, "--bg-color", "FF0000")
	assertExit(t, r, 2) // missing --table-index
}

// ---------------------------------------------------------------------------
// Default Font E2E
// ---------------------------------------------------------------------------

func TestExcel_DefaultFont_Get(t *testing.T) {
	env := setupHome(t, "write")
	file := filepath.Join(t.TempDir(), "font.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["A","B"],["C","D"]]}]}`
	r0 := run(t, env, "excel", "create", file, "--spec", spec, "--json")
	assertExit(t, r0, 0)

	r := run(t, env, "excel", "default-font", file, "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["defaultFont"] == nil || m["defaultFont"] == "" {
		t.Error("expected non-empty defaultFont")
	}
}

func TestExcel_DefaultFont_Set(t *testing.T) {
	env := setupHome(t, "write")
	file := filepath.Join(t.TempDir(), "font.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["A","B"],["C","D"]]}]}`
	r0 := run(t, env, "excel", "create", file, "--spec", spec, "--json")
	assertExit(t, r0, 0)

	// Set
	r := run(t, env, "excel", "default-font", file, "--set", "Arial", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["status"] != "ok" {
		t.Errorf("status=%v want ok", m["status"])
	}

	// Verify
	r2 := run(t, env, "excel", "default-font", file, "--json")
	assertExit(t, r2, 0)
	var m2 map[string]any
	parseJSON(t, r2.Stdout, &m2)
	if m2["defaultFont"] != "Arial" {
		t.Errorf("defaultFont=%v want Arial", m2["defaultFont"])
	}
}

// ---------------------------------------------------------------------------
// Cell Style E2E
// ---------------------------------------------------------------------------

func TestExcel_CellStyle_Read(t *testing.T) {
	env := setupHome(t, "write")
	file := filepath.Join(t.TempDir(), "style.xlsx")
	run(t, env, "excel", "create", file, "--spec", `{"sheets":[{"name":"S","rows":[["Name","Value"],["Alice",100]]}]}`, "--json")

	// Apply a style
	run(t, env, "excel", "style", file, "--sheet", "S", "--range", "A1:B1",
		"--bold", "--font-size", "14", "--font-family", "Arial", "--json")

	// Read it back
	r := run(t, env, "excel", "cell-style", file, "A1", "--sheet", "S", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	if m["bold"] != true {
		t.Errorf("bold=%v want true", m["bold"])
	}
	if fontSize, ok := m["fontSize"].(float64); !ok || fontSize != 14 {
		t.Errorf("fontSize=%v want 14", m["fontSize"])
	}
	if m["fontFamily"] != "Arial" {
		t.Errorf("fontFamily=%v want Arial", m["fontFamily"])
	}
}

func TestExcel_CellStyle_Default(t *testing.T) {
	env := setupHome(t, "write")
	file := filepath.Join(t.TempDir(), "style.xlsx")
	spec := `{"sheets":[{"name":"S","rows":[["A","B"],["C","D"]]}]}`
	r0 := run(t, env, "excel", "create", file, "--spec", spec, "--json")
	assertExit(t, r0, 0)

	// An unstyled cell should have style index 0 and no bold.
	r := run(t, env, "excel", "cell-style", file, "A1", "--sheet", "S", "--json")
	assertExit(t, r, 0)
	var m map[string]any
	parseJSON(t, r.Stdout, &m)
	// styleIndex 0 means default — bold should be false/absent.
	if m["bold"] == true {
		t.Error("unstyled cell should not be bold")
	}
}
