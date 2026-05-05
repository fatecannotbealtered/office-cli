package excel

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// makeXLSX is a tiny helper that creates an .xlsx with one sheet and a few rows.
// Tests use this rather than a checked-in fixture so failures point at the test
// inputs directly.
func makeXLSX(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "sample.xlsx")

	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	first := f.GetSheetName(0)
	if first != "Sales" {
		if _, err := f.NewSheet("Sales"); err != nil {
			t.Fatal(err)
		}
		if err := f.DeleteSheet(first); err != nil {
			t.Fatal(err)
		}
	}
	rows := [][]any{{"Name", "Sales"}, {"Alice", 100}, {"Bob", 200}}
	for r, row := range rows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellValue("Sales", ref, val); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestReadTyped(t *testing.T) {
	path := makeXLSX(t)

	res, err := ReadTyped(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("ReadTyped: %v", err)
	}
	if res.Sheet != "Sales" {
		t.Errorf("sheet=%q want Sales", res.Sheet)
	}
	if len(res.Headers) != 2 || res.Headers[0] != "Name" || res.Headers[1] != "Sales" {
		t.Errorf("headers=%v want [Name Sales]", res.Headers)
	}
	// Types: Name→string, Sales→number
	if res.Types["Name"] != "string" {
		t.Errorf("Name type=%q want string", res.Types["Name"])
	}
	if res.Types["Sales"] != "number" {
		t.Errorf("Sales type=%q want number", res.Types["Sales"])
	}
	// Rows: first data row is Alice/100
	if len(res.Rows) != 2 {
		t.Fatalf("rows=%d want 2", len(res.Rows))
	}
	if res.Rows[0]["Name"] != "Alice" {
		t.Errorf("row0 Name=%v want Alice", res.Rows[0]["Name"])
	}
	// Number should be int (100), not string "100"
	if res.Rows[0]["Sales"] != 100 {
		t.Errorf("row0 Sales=%v (%T) want int 100", res.Rows[0]["Sales"], res.Rows[0]["Sales"])
	}
}

func TestReadTypedWithRange(t *testing.T) {
	path := makeXLSX(t)

	res, err := ReadTyped(path, ReadOptions{Sheet: "Sales", Range: "A1:B2"})
	if err != nil {
		t.Fatalf("ReadTyped range: %v", err)
	}
	if len(res.Rows) != 1 {
		t.Errorf("rows=%d want 1", len(res.Rows))
	}
	if res.Rows[0]["Name"] != "Alice" {
		t.Errorf("row0 Name=%v want Alice", res.Rows[0]["Name"])
	}
}

func TestListSheetsAndRead(t *testing.T) {
	path := makeXLSX(t)

	sheets, err := ListSheets(path)
	if err != nil {
		t.Fatalf("ListSheets: %v", err)
	}
	if len(sheets) == 0 || sheets[0].Name != "Sales" {
		t.Fatalf("unexpected sheets: %+v", sheets)
	}

	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if len(res.Rows) != 3 || res.Rows[1][0] != "Alice" {
		t.Fatalf("unexpected rows: %+v", res.Rows)
	}
}

func TestReadCellAndSearch(t *testing.T) {
	path := makeXLSX(t)

	val, _, err := ReadCell(path, "Sales", "B2")
	if err != nil {
		t.Fatalf("ReadCell: %v", err)
	}
	if val != "100" {
		t.Errorf("expected 100, got %q", val)
	}

	matches, err := Search(path, "Bob", "", 10)
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	if len(matches) != 1 || matches[0].Ref != "A3" {
		t.Errorf("unexpected matches: %+v", matches)
	}
}

func TestAppendRows(t *testing.T) {
	path := makeXLSX(t)

	startRow, err := AppendRows(path, "Sales", [][]any{{"Carol", 300}})
	if err != nil {
		t.Fatalf("AppendRows: %v", err)
	}
	// AppendRows returns the 1-based row index where data started being written.
	if startRow != 4 {
		t.Errorf("expected append to start at row 4, got %d", startRow)
	}
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after append: %v", err)
	}
	last := res.Rows[len(res.Rows)-1]
	if last[0] != "Carol" || last[1] != "300" {
		t.Errorf("expected Carol/300 at end, got %+v", last)
	}
}

func TestSheetLifecycle(t *testing.T) {
	path := makeXLSX(t)

	if _, err := CopySheet(path, "Sales", "Sales-copy"); err != nil {
		t.Fatalf("CopySheet: %v", err)
	}
	if err := RenameSheet(path, "Sales-copy", "Backup"); err != nil {
		t.Fatalf("RenameSheet: %v", err)
	}
	sheets, _ := ListSheets(path)
	names := map[string]bool{}
	for _, s := range sheets {
		names[s.Name] = true
	}
	if !names["Backup"] {
		t.Errorf("Backup sheet missing: %v", names)
	}
	if err := DeleteSheet(path, "Backup"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}
}

func TestToCSVAndFromCSV(t *testing.T) {
	path := makeXLSX(t)
	dir := t.TempDir()
	csv := filepath.Join(dir, "out.csv")

	if err := ToCSV(path, "Sales", csv, true); err != nil {
		t.Fatalf("ToCSV: %v", err)
	}

	xlsx := filepath.Join(dir, "round.xlsx")
	if _, err := FromCSV(csv, xlsx, "csv", true); err != nil {
		t.Fatalf("FromCSV: %v", err)
	}
	res, err := Read(xlsx, ReadOptions{Sheet: "csv"})
	if err != nil {
		t.Fatalf("Read round-trip: %v", err)
	}
	if len(res.Rows) < 3 {
		t.Errorf("round trip lost rows: %+v", res.Rows)
	}
}

// ---------------------------------------------------------------------------
// Feature 2: Formula generation
// ---------------------------------------------------------------------------

func TestGenerateFormula(t *testing.T) {
	path := makeXLSX(t) // Sales sheet: row1=headers, row2=Alice/100, row3=Bob/200

	res, err := GenerateFormula(path, "Sales", FormulaSpec{Column: "B", Func: "SUM"})
	if err != nil {
		t.Fatalf("GenerateFormula: %v", err)
	}
	// Data range should auto-detect B2:B3 (skip header row 1)
	if res.DataRange != "B2:B3" {
		t.Errorf("DataRange=%q want B2:B3", res.DataRange)
	}
	if res.Formula != "=SUM(B2:B3)" {
		t.Errorf("Formula=%q want =SUM(B2:B3)", res.Formula)
	}
	// Output cell should be one row below the data: B4
	if res.Cell != "B4" {
		t.Errorf("Cell=%q want B4", res.Cell)
	}
	if res.DataRows != 2 {
		t.Errorf("DataRows=%d want 2", res.DataRows)
	}
}

func TestGenerateFormulaAverage(t *testing.T) {
	path := makeXLSX(t)

	res, err := GenerateFormula(path, "Sales", FormulaSpec{Column: "B", Func: "AVERAGE"})
	if err != nil {
		t.Fatalf("GenerateFormula AVERAGE: %v", err)
	}
	if res.Formula != "=AVERAGE(B2:B3)" {
		t.Errorf("Formula=%q want =AVERAGE(B2:B3)", res.Formula)
	}
}

func TestGenerateFormulaWithCriteria(t *testing.T) {
	path := makeXLSX(t)

	res, err := GenerateFormula(path, "Sales", FormulaSpec{
		Column:   "B",
		Func:     "COUNTIF",
		Criteria: ">100",
	})
	if err != nil {
		t.Fatalf("GenerateFormula COUNTIF: %v", err)
	}
	if res.Formula != `=COUNTIF(B2:B3,">100")` {
		t.Errorf("Formula=%q want =COUNTIF(B2:B3,\">100\")", res.Formula)
	}
}

func TestGenerateFormulaExplicitRange(t *testing.T) {
	path := makeXLSX(t)

	res, err := GenerateFormula(path, "Sales", FormulaSpec{
		Column:    "B",
		Func:      "SUM",
		DataRange: "B2:B2",
	})
	if err != nil {
		t.Fatalf("GenerateFormula explicit range: %v", err)
	}
	if res.DataRange != "B2:B2" {
		t.Errorf("DataRange=%q want B2:B2", res.DataRange)
	}
	if res.DataRows != 1 {
		t.Errorf("DataRows=%d want 1", res.DataRows)
	}
}

func TestGenerateFormulasBatch(t *testing.T) {
	path := makeXLSX(t)

	results, err := GenerateFormulas(path, "Sales", []FormulaSpec{
		{Column: "B", Func: "SUM"},
		{Column: "B", Func: "AVERAGE"},
	})
	if err != nil {
		t.Fatalf("GenerateFormulas: %v", err)
	}
	if len(results) != 2 {
		t.Fatalf("results=%d want 2", len(results))
	}
	if results[0].Formula != "=SUM(B2:B3)" {
		t.Errorf("results[0].Formula=%q", results[0].Formula)
	}
	if results[1].Formula != "=AVERAGE(B2:B3)" {
		t.Errorf("results[1].Formula=%q", results[1].Formula)
	}
}

func TestApplyFormulas(t *testing.T) {
	path := makeXLSX(t)

	results := []FormulaResult{
		{Cell: "B4", Formula: "=SUM(B2:B3)", DataRows: 2, DataRange: "B2:B3"},
	}
	if err := ApplyFormulas(path, "Sales", results); err != nil {
		t.Fatalf("ApplyFormulas: %v", err)
	}
	// Read back the cell to verify the formula was written.
	val, formula, err := ReadCell(path, "Sales", "B4")
	if err != nil {
		t.Fatalf("ReadCell after ApplyFormulas: %v", err)
	}
	if formula != "=SUM(B2:B3)" {
		t.Errorf("formula=%q want =SUM(B2:B3)", formula)
	}
	_ = val // value will be the cached result, not important here
}

// ---------------------------------------------------------------------------
// Feature 3: Copy workbook
// ---------------------------------------------------------------------------

func TestCopyWorkbook(t *testing.T) {
	src := makeXLSX(t)
	dir := t.TempDir()
	dst := filepath.Join(dir, "copy.xlsx")

	if err := CopyWorkbook(src, dst, false); err != nil {
		t.Fatalf("CopyWorkbook: %v", err)
	}
	// Verify the copy can be read and has the same data.
	res, err := Read(dst, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read copy: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3", len(res.Rows))
	}
	if res.Rows[1][0] != "Alice" || res.Rows[1][1] != "100" {
		t.Errorf("row1=%v want [Alice 100]", res.Rows[1])
	}
}

func TestCopyWorkbookNoOverwrite(t *testing.T) {
	src := makeXLSX(t)
	dir := t.TempDir()
	dst := filepath.Join(dir, "copy.xlsx")

	// Create the destination file first.
	if err := CopyWorkbook(src, dst, true); err != nil {
		t.Fatalf("first copy: %v", err)
	}
	// Second copy without overwrite should fail.
	if err := CopyWorkbook(src, dst, false); err == nil {
		t.Fatal("expected error when copying without overwrite, got nil")
	}
	// With overwrite should succeed.
	if err := CopyWorkbook(src, dst, true); err != nil {
		t.Fatalf("overwrite copy: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Feature 4: Data validation
// ---------------------------------------------------------------------------

func TestAddValidationDropdown(t *testing.T) {
	path := makeXLSX(t)

	err := AddValidation(path, "Sales", ValidationSpec{
		Type:  "list",
		Range: "A2:A100",
		List:  []string{"Alice", "Bob", "Carol"},
	})
	if err != nil {
		t.Fatalf("AddValidation dropdown: %v", err)
	}
	// We can't easily read back validation rules with excelize, but we can
	// verify the file still opens and reads correctly.
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after validation: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3 (data should be intact)", len(res.Rows))
	}
}

func TestAddValidationNumberRange(t *testing.T) {
	path := makeXLSX(t)

	err := AddValidation(path, "Sales", ValidationSpec{
		Type:     "whole",
		Range:    "B2:B100",
		Operator: "between",
		Min:      "0",
		Max:      "1000",
	})
	if err != nil {
		t.Fatalf("AddValidation number range: %v", err)
	}
}

func TestAddValidationCustom(t *testing.T) {
	path := makeXLSX(t)

	err := AddValidation(path, "Sales", ValidationSpec{
		Type:    "custom",
		Range:   "A2:A100",
		Formula: "=LEN(A2)>0",
	})
	if err != nil {
		t.Fatalf("AddValidation custom: %v", err)
	}
}

func TestAddValidationWithMessages(t *testing.T) {
	path := makeXLSX(t)

	err := AddValidation(path, "Sales", ValidationSpec{
		Type:       "whole",
		Range:      "B2:B100",
		Min:        "0",
		Max:        "100",
		ErrorMsg:   "Must be 0-100",
		ErrorTitle: "Invalid",
		PromptMsg:  "Enter a number",
	})
	if err != nil {
		t.Fatalf("AddValidation with messages: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Feature 5: Batch style
// ---------------------------------------------------------------------------

func TestStyleCellsBatch(t *testing.T) {
	path := makeXLSX(t)

	err := StyleCellsBatch(path, "Sales", []BatchStyleEntry{
		{
			Range: "A1:B1",
			StyleSpec: StyleSpec{
				Bold:      true,
				BgColor:   "4472C4",
				TextColor: "FFFFFF",
				Align:     "center",
			},
		},
		{
			Range: "A2:B3",
			StyleSpec: StyleSpec{
				NumberFormat: "#,##0.00",
			},
		},
	})
	if err != nil {
		t.Fatalf("StyleCellsBatch: %v", err)
	}
	// Verify the file still reads correctly.
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after batch style: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3", len(res.Rows))
	}
}

func TestStyleCellsBatchEmpty(t *testing.T) {
	path := makeXLSX(t)

	err := StyleCellsBatch(path, "Sales", []BatchStyleEntry{})
	if err == nil {
		t.Fatal("expected error for empty entries, got nil")
	}
}

// ---------------------------------------------------------------------------
// Feature 2b: Column-level formula generation
// ---------------------------------------------------------------------------

func TestColumnFormulaVLOOKUP(t *testing.T) {
	path := makeXLSX(t) // Sales: Name(A), Sales(B)

	spec := ColumnFormulaSpec{
		Column:     "C",
		Func:       "VLOOKUP",
		LookupCol:  "A",
		ReturnCol:  "B",
		TableRange: "A2:B3",
	}
	result, err := GenerateColumnFormula(path, "Sales", spec)
	if err != nil {
		t.Fatalf("GenerateColumnFormula VLOOKUP: %v", err)
	}
	expected := "=VLOOKUP(A2,A2:B3,2,0)"
	if result.Formula != expected {
		t.Errorf("formula=%q want %q", result.Formula, expected)
	}
	if result.RowCount != 2 {
		t.Errorf("rowCount=%d want 2", result.RowCount)
	}
}

func TestColumnFormulaIF(t *testing.T) {
	path := makeXLSX(t)

	spec := ColumnFormulaSpec{
		Column:       "C",
		Func:         "IF",
		Condition:    "B{row}>150",
		ValueIfTrue:  "\"High\"",
		ValueIfFalse: "\"Low\"",
	}
	result, err := GenerateColumnFormula(path, "Sales", spec)
	if err != nil {
		t.Fatalf("GenerateColumnFormula IF: %v", err)
	}
	expected := `=IF(B2>150,"High","Low")`
	if result.Formula != expected {
		t.Errorf("formula=%q want %q", result.Formula, expected)
	}
}

func TestColumnFormulaCONCAT(t *testing.T) {
	path := makeXLSX(t)

	spec := ColumnFormulaSpec{
		Column:     "C",
		Func:       "CONCAT",
		ConcatCols: []string{"A", "B"},
		Separator:  " - ",
	}
	result, err := GenerateColumnFormula(path, "Sales", spec)
	if err != nil {
		t.Fatalf("GenerateColumnFormula CONCAT: %v", err)
	}
	expected := `=CONCATENATE(A2," - ",B2)`
	if result.Formula != expected {
		t.Errorf("formula=%q want %q", result.Formula, expected)
	}
}

func TestColumnFormulaCUSTOM(t *testing.T) {
	path := makeXLSX(t)

	spec := ColumnFormulaSpec{
		Column:        "C",
		Func:          "CUSTOM",
		CustomFormula: "=B{row}*2",
	}
	result, err := GenerateColumnFormula(path, "Sales", spec)
	if err != nil {
		t.Fatalf("GenerateColumnFormula CUSTOM: %v", err)
	}
	if result.Formula != "=B2*2" {
		t.Errorf("formula=%q want =B2*2", result.Formula)
	}
}

func TestApplyColumnFormulas(t *testing.T) {
	path := makeXLSX(t)

	spec := ColumnFormulaSpec{
		Column:        "C",
		Func:          "CUSTOM",
		CustomFormula: "=B{row}*2",
	}
	count, err := ApplyColumnFormulas(path, "Sales", spec)
	if err != nil {
		t.Fatalf("ApplyColumnFormulas: %v", err)
	}
	if count != 2 {
		t.Errorf("count=%d want 2", count)
	}
	// Verify the formula was written.
	_, formula, err := ReadCell(path, "Sales", "C2")
	if err != nil {
		t.Fatalf("ReadCell: %v", err)
	}
	if formula != "=B2*2" {
		t.Errorf("formula=%q want =B2*2", formula)
	}
}

// ---------------------------------------------------------------------------
// Feature 6: Excel ↔ JSON
// ---------------------------------------------------------------------------

func TestToJSON(t *testing.T) {
	path := makeXLSX(t)

	// With headers, typed
	data, err := ToJSON(path, JSONReadOptions{Sheet: "Sales", WithHeaders: true, Typed: true})
	if err != nil {
		t.Fatalf("ToJSON: %v", err)
	}
	rows, ok := data.([]map[string]any)
	if !ok {
		t.Fatalf("expected []map[string]any, got %T", data)
	}
	if len(rows) != 2 {
		t.Errorf("rows=%d want 2", len(rows))
	}
	if rows[0]["Name"] != "Alice" {
		t.Errorf("row0 Name=%v want Alice", rows[0]["Name"])
	}
	if rows[0]["Sales"] != 100 {
		t.Errorf("row0 Sales=%v want 100", rows[0]["Sales"])
	}
}

func TestToJSONString(t *testing.T) {
	path := makeXLSX(t)

	data, err := ToJSON(path, JSONReadOptions{Sheet: "Sales", WithHeaders: true})
	if err != nil {
		t.Fatalf("ToJSON string: %v", err)
	}
	rows, ok := data.([]map[string]string)
	if !ok {
		t.Fatalf("expected []map[string]string, got %T", data)
	}
	if len(rows) != 2 {
		t.Errorf("rows=%d want 2", len(rows))
	}
}

func TestFromJSONObjects(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "fromjson.xlsx")

	data := []map[string]any{
		{"Name": "Alice", "Score": 95},
		{"Name": "Bob", "Score": 80},
	}
	rows, err := FromJSON(path, "Data", data, false)
	if err != nil {
		t.Fatalf("FromJSON: %v", err)
	}
	if rows != 3 { // 1 header + 2 data
		t.Errorf("rows=%d want 3", rows)
	}
	// Read back.
	res, err := Read(path, ReadOptions{Sheet: "Data"})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("read rows=%d want 3", len(res.Rows))
	}
}

func TestFromJSONArrays(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "fromjson_arr.xlsx")

	data := [][]any{
		{"A", "B", "C"},
		{1, 2, 3},
	}
	rows, err := FromJSON(path, "Sheet1", data, false)
	if err != nil {
		t.Fatalf("FromJSON arrays: %v", err)
	}
	if rows != 2 {
		t.Errorf("rows=%d want 2", rows)
	}
}

// ---------------------------------------------------------------------------
// Feature 7: Range operations
// ---------------------------------------------------------------------------

func TestFillRange(t *testing.T) {
	path := makeXLSX(t)

	if err := FillRange(path, "Sales", "C1:C5", "N/A"); err != nil {
		t.Fatalf("FillRange: %v", err)
	}
	// Verify.
	val, _, err := ReadCell(path, "Sales", "C3")
	if err != nil {
		t.Fatalf("ReadCell: %v", err)
	}
	if val != "N/A" {
		t.Errorf("value=%q want N/A", val)
	}
}

func TestCopyRange(t *testing.T) {
	path := makeXLSX(t)

	// Copy A1:B3 to C1.
	if err := CopyRange(path, "Sales", "A1:B3", "C1"); err != nil {
		t.Fatalf("CopyRange: %v", err)
	}
	// Verify C1 = "Name" (copied from A1).
	val, _, err := ReadCell(path, "Sales", "C1")
	if err != nil {
		t.Fatalf("ReadCell C1: %v", err)
	}
	if val != "Name" {
		t.Errorf("C1=%q want Name", val)
	}
	// Verify C2 = "Alice" (copied from A2).
	val, _, err = ReadCell(path, "Sales", "C2")
	if err != nil {
		t.Fatalf("ReadCell C2: %v", err)
	}
	if val != "Alice" {
		t.Errorf("C2=%q want Alice", val)
	}
}

// ---------------------------------------------------------------------------
// Feature 8: Multi-series chart
// ---------------------------------------------------------------------------

func TestAddMultiChart(t *testing.T) {
	// Create a sheet with multiple columns.
	dir := t.TempDir()
	path := filepath.Join(dir, "chart.xlsx")
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	sheet := f.GetSheetName(0)
	rows := [][]any{
		{"Month", "Sales", "Profit"},
		{"Jan", 100, 30},
		{"Feb", 200, 60},
		{"Mar", 150, 45},
	}
	for r, row := range rows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
			_ = f.SetCellValue(sheet, ref, val)
		}
	}
	_ = f.SaveAs(path)

	err := AddMultiChart(path, sheet, MultiChartSpec{
		Type:  "line",
		Cell:  "E2",
		Title: "Revenue",
		Series: []SeriesSpec{
			{Name: "Sales", CatRange: "A2:A4", ValRange: "B2:B4"},
			{Name: "Profit", CatRange: "A2:A4", ValRange: "C2:C4"},
		},
	})
	if err != nil {
		t.Fatalf("AddMultiChart: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Feature 9: Data bar and icon set
// ---------------------------------------------------------------------------

func TestAddDataBar(t *testing.T) {
	path := makeXLSX(t)

	err := AddDataBar(path, "Sales", DataBarSpec{
		Range:    "B2:B3",
		MinColor: "F8696B",
		MaxColor: "63BE7B",
	})
	if err != nil {
		t.Fatalf("AddDataBar: %v", err)
	}
	// Verify file still reads.
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after data bar: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3", len(res.Rows))
	}
}

func TestAddIconSet(t *testing.T) {
	path := makeXLSX(t)

	err := AddIconSet(path, "Sales", IconSetSpec{
		Range: "B2:B3",
		Style: "3TrafficLights1",
	})
	if err != nil {
		t.Fatalf("AddIconSet: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Feature 10: Color scale conditional formatting
// ---------------------------------------------------------------------------

func TestAddColorScale2(t *testing.T) {
	path := makeXLSX(t)

	err := AddColorScale(path, "Sales", ColorScaleSpec{
		Range:    "B2:B3",
		MinColor: "F8696B",
		MaxColor: "63BE7B",
	})
	if err != nil {
		t.Fatalf("AddColorScale 2-color: %v", err)
	}
	// Verify file still reads.
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after color scale: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3", len(res.Rows))
	}
}

func TestAddColorScale3(t *testing.T) {
	path := makeXLSX(t)

	err := AddColorScale(path, "Sales", ColorScaleSpec{
		Range:    "B2:B3",
		MinColor: "F8696B",
		MidColor: "FFEB84",
		MaxColor: "63BE7B",
	})
	if err != nil {
		t.Fatalf("AddColorScale 3-color: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Feature 11: Hyperlinks
// ---------------------------------------------------------------------------

func TestSetAndGetHyperlink(t *testing.T) {
	path := makeXLSX(t)

	err := SetHyperlink(path, "Sales", HyperlinkSpec{
		Cell:    "A2",
		Link:    "https://example.com",
		Display: "Alice (click)",
		Tooltip: "Open Alice's profile",
	})
	if err != nil {
		t.Fatalf("SetHyperlink: %v", err)
	}

	// Read back.
	res, err := GetHyperlink(path, "Sales", "A2")
	if err != nil {
		t.Fatalf("GetHyperlink: %v", err)
	}
	if !res.Exists {
		t.Error("expected hyperlink to exist")
	}
	if res.Link != "https://example.com" {
		t.Errorf("link=%q want https://example.com", res.Link)
	}
}

func TestSetHyperlinkInternal(t *testing.T) {
	path := makeXLSX(t)

	err := SetHyperlink(path, "Sales", HyperlinkSpec{
		Cell: "A3",
		Link: "Sales!A1",
		Type: "Location",
	})
	if err != nil {
		t.Fatalf("SetHyperlink internal: %v", err)
	}

	res, err := GetHyperlink(path, "Sales", "A3")
	if err != nil {
		t.Fatalf("GetHyperlink: %v", err)
	}
	if !res.Exists {
		t.Error("expected hyperlink to exist")
	}
}

func TestGetHyperlinkNone(t *testing.T) {
	path := makeXLSX(t)

	res, err := GetHyperlink(path, "Sales", "A1")
	if err != nil {
		t.Fatalf("GetHyperlink: %v", err)
	}
	if res.Exists {
		t.Error("expected no hyperlink on A1")
	}
}

// ---------------------------------------------------------------------------
// StyleSpec: new fields (strike, textRotation, indent, shrinkToFit)
// ---------------------------------------------------------------------------

func TestStyleStrikeAndRotation(t *testing.T) {
	path := makeXLSX(t)

	err := StyleCells(path, "Sales", "A1:B1", StyleSpec{
		Strike:       true,
		TextRotation: 45,
	})
	if err != nil {
		t.Fatalf("StyleCells strike/rotation: %v", err)
	}
	// Verify file still reads.
	res, err := Read(path, ReadOptions{Sheet: "Sales"})
	if err != nil {
		t.Fatalf("Read after style: %v", err)
	}
	if len(res.Rows) != 3 {
		t.Errorf("rows=%d want 3", len(res.Rows))
	}
}

func TestStyleIndentAndShrink(t *testing.T) {
	path := makeXLSX(t)

	err := StyleCells(path, "Sales", "A2:A3", StyleSpec{
		Indent:      2,
		ShrinkToFit: true,
	})
	if err != nil {
		t.Fatalf("StyleCells indent/shrink: %v", err)
	}
}

// ---------------------------------------------------------------------------
// Default Font Tests
// ---------------------------------------------------------------------------

func TestGetDefaultFont(t *testing.T) {
	path := makeXLSX(t)
	name, err := GetDefaultFont(path)
	if err != nil {
		t.Fatalf("GetDefaultFont: %v", err)
	}
	// excelize default is "Calibri".
	if name == "" {
		t.Error("expected non-empty default font name")
	}
}

func TestSetAndGetDefaultFont(t *testing.T) {
	path := makeXLSX(t)
	err := SetDefaultFont(path, "Arial")
	if err != nil {
		t.Fatalf("SetDefaultFont: %v", err)
	}
	name, err := GetDefaultFont(path)
	if err != nil {
		t.Fatalf("GetDefaultFont after set: %v", err)
	}
	if name != "Arial" {
		t.Errorf("font=%q want Arial", name)
	}
}

func TestSetDefaultFontEmpty(t *testing.T) {
	path := makeXLSX(t)
	err := SetDefaultFont(path, "")
	if err == nil {
		t.Error("expected error for empty font name")
	}
}

// ---------------------------------------------------------------------------
// Cell Style Read Tests
// ---------------------------------------------------------------------------

func TestGetCellStyleInfoDefault(t *testing.T) {
	path := makeXLSX(t)
	// An unstyled cell should return style index 0 (default).
	info, err := GetCellStyleInfo(path, "Sales", "A2")
	if err != nil {
		t.Fatalf("GetCellStyleInfo: %v", err)
	}
	if info.Cell != "A2" {
		t.Errorf("cell=%q want A2", info.Cell)
	}
	// Default style index is 0 — should not have bold.
	if info.Bold {
		t.Error("unstyled cell should not be bold")
	}
}

func TestGetCellStyleInfoStyled(t *testing.T) {
	path := makeXLSX(t)
	// Apply a style.
	err := StyleCells(path, "Sales", "A1:B1", StyleSpec{
		Bold:         true,
		Italic:       true,
		FontSize:     14,
		FontFamily:   "Arial",
		TextColor:    "FF0000",
		BgColor:      "0000FF",
		NumberFormat: "#,##0.00",
		Align:        "center",
		Valign:       "center",
		WrapText:     true,
	})
	if err != nil {
		t.Fatalf("StyleCells: %v", err)
	}
	// Read it back.
	info, err := GetCellStyleInfo(path, "Sales", "A1")
	if err != nil {
		t.Fatalf("GetCellStyleInfo: %v", err)
	}
	if !info.Bold {
		t.Error("expected bold")
	}
	if !info.Italic {
		t.Error("expected italic")
	}
	if info.FontSize != 14 {
		t.Errorf("fontSize=%g want 14", info.FontSize)
	}
	if info.FontFamily != "Arial" {
		t.Errorf("fontFamily=%q want Arial", info.FontFamily)
	}
	if info.FontColor != "FF0000" {
		t.Errorf("fontColor=%q want FF0000", info.FontColor)
	}
	if info.BgColor != "0000FF" {
		t.Errorf("bgColor=%q want 0000FF", info.BgColor)
	}
	if info.Align != "center" {
		t.Errorf("align=%q want center", info.Align)
	}
	if info.Valign != "center" {
		t.Errorf("valign=%q want center", info.Valign)
	}
	if !info.WrapText {
		t.Error("expected wrapText")
	}
	if info.StyleIndex == 0 {
		t.Error("styled cell should not have style index 0")
	}
}

func TestGetCellStyleInfoStrikeAndRotation(t *testing.T) {
	path := makeXLSX(t)
	err := StyleCells(path, "Sales", "A1:A1", StyleSpec{
		Strike:       true,
		TextRotation: 45,
		Indent:       3,
		ShrinkToFit:  true,
	})
	if err != nil {
		t.Fatalf("StyleCells: %v", err)
	}
	info, err := GetCellStyleInfo(path, "Sales", "A1")
	if err != nil {
		t.Fatalf("GetCellStyleInfo: %v", err)
	}
	if !info.Strike {
		t.Error("expected strike")
	}
	if info.TextRotation != 45 {
		t.Errorf("textRotation=%d want 45", info.TextRotation)
	}
	if info.Indent != 3 {
		t.Errorf("indent=%d want 3", info.Indent)
	}
	if !info.ShrinkToFit {
		t.Error("expected shrinkToFit")
	}
}
