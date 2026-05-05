// Package excel wraps excelize and exposes a thin, AI-friendly API for
// reading, writing, searching and converting xlsx workbooks.
//
// All functions operate on file paths (rather than open handles) so callers
// don't have to track lifetimes. Internally we open, mutate, save and close.
package excel

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/xuri/excelize/v2"
)

// SheetInfo summarizes one worksheet for `excel sheets`.
type SheetInfo struct {
	Name      string `json:"name"`
	Index     int    `json:"index"`
	Rows      int    `json:"rows"`
	Cols      int    `json:"cols"`
	Hidden    bool   `json:"hidden,omitempty"`
	Dimension string `json:"dimension,omitempty"`
}

// SheetMatch is one search hit returned by Search.
type SheetMatch struct {
	Sheet   string `json:"sheet"`
	Ref     string `json:"ref"`
	Value   string `json:"value"`
	Formula string `json:"formula,omitempty"`
}

// CreateSpec describes a new workbook for Create. AI Agents pass this as JSON.
type CreateSpec struct {
	Sheets []SheetSpec `json:"sheets"`
}

// SheetSpec is one sheet inside a CreateSpec.
type SheetSpec struct {
	Name string     `json:"name"`
	Rows [][]any    `json:"rows,omitempty"`
	Cols []ColStyle `json:"cols,omitempty"`
}

// ColStyle is a minimal column style descriptor (width only, for now).
type ColStyle struct {
	Width float64 `json:"width,omitempty"`
}

// open is the canonical entry point. Returns a closer the caller MUST defer.
func open(path string) (*excelize.File, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return f, nil
}

// ListSheets returns one entry per worksheet, including hidden ones.
func ListSheets(path string) ([]SheetInfo, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	names := f.GetSheetList()
	out := make([]SheetInfo, 0, len(names))
	for i, name := range names {
		info := SheetInfo{Name: name, Index: i}
		dim, _ := f.GetSheetDimension(name)
		info.Dimension = dim
		rows, _ := f.GetRows(name)
		info.Rows = len(rows)
		if len(rows) > 0 {
			maxCols := 0
			for _, row := range rows {
				if len(row) > maxCols {
					maxCols = len(row)
				}
			}
			info.Cols = maxCols
		}
		visibility, _ := f.GetSheetVisible(name)
		info.Hidden = !visibility
		out = append(out, info)
	}
	return out, nil
}

// ReadResult is the response shape of Read.
type ReadResult struct {
	Sheet string     `json:"sheet"`
	Range string     `json:"range,omitempty"`
	Rows  [][]string `json:"rows"`
}

// ReadOptions narrows down what Read returns.
type ReadOptions struct {
	Sheet string // sheet name OR empty for first sheet
	Range string // "A1:D10" or "" for full sheet
	Limit int    // max rows; 0 = all
}

// Read returns the rows of one sheet, optionally limited by a range and a row cap.
func Read(path string, opts ReadOptions) (*ReadResult, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	sheet := opts.Sheet
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return nil, fmt.Errorf("sheet not found: %s", sheet)
	}

	var rows [][]string
	if opts.Range != "" {
		rows, err = readRange(f, sheet, opts.Range)
	} else {
		rows, err = f.GetRows(sheet)
	}
	if err != nil {
		return nil, err
	}

	if opts.Limit > 0 && len(rows) > opts.Limit {
		rows = rows[:opts.Limit]
	}

	return &ReadResult{Sheet: sheet, Range: opts.Range, Rows: rows}, nil
}

// ---------------------------------------------------------------------------
// Typed read — returns cell type information alongside values
// ---------------------------------------------------------------------------

// TypedReadResult is the response shape of ReadTyped.
type TypedReadResult struct {
	Sheet   string            `json:"sheet"`
	Range   string            `json:"range,omitempty"`
	Headers []string          `json:"headers,omitempty"`
	Rows    []map[string]any  `json:"rows"`
	Types   map[string]string `json:"types"`
}

// ReadTyped returns cell values with their native Go types and type annotations.
// The first row is always treated as headers. Each data row is a map keyed by
// header name, with values converted to their natural Go types (float64, bool,
// string). The Types map records the Excel-level type of each header column.
func ReadTyped(path string, opts ReadOptions) (*TypedReadResult, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	sheet := opts.Sheet
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return nil, fmt.Errorf("sheet not found: %s", sheet)
	}

	// Determine the range to read.
	var startCol, startRow, endCol, endRow int
	if opts.Range != "" {
		parts := strings.Split(opts.Range, ":")
		if len(parts) != 2 {
			return nil, fmt.Errorf("invalid range: %s (expected e.g. A1:D10)", opts.Range)
		}
		startCol, startRow, err = excelize.CellNameToCoordinates(strings.TrimSpace(parts[0]))
		if err != nil {
			return nil, fmt.Errorf("invalid start cell: %w", err)
		}
		endCol, endRow, err = excelize.CellNameToCoordinates(strings.TrimSpace(parts[1]))
		if err != nil {
			return nil, fmt.Errorf("invalid end cell: %w", err)
		}
	} else {
		startCol, startRow = 1, 1
		// Detect dimensions from the sheet.
		dim, _ := f.GetSheetDimension(sheet)
		if dim != "" {
			parts := strings.Split(dim, ":")
			if len(parts) == 2 {
				endCol, endRow, _ = excelize.CellNameToCoordinates(strings.TrimSpace(parts[1]))
			}
		}
		if endCol == 0 || endRow == 0 {
			// Fallback: use GetRows to find dimensions.
			rows, _ := f.GetRows(sheet)
			endRow = len(rows)
			if endRow > 0 {
				maxCols := 0
				for _, row := range rows {
					if len(row) > maxCols {
						maxCols = len(row)
					}
				}
				endCol = maxCols
			}
		}
		if endCol == 0 {
			endCol = 1
		}
		if endRow == 0 {
			endRow = 1
		}
	}

	// Read headers from the first row of the range.
	var headers []string
	for c := startCol; c <= endCol; c++ {
		ref, _ := excelize.CoordinatesToCellName(c, startRow)
		val, _ := f.GetCellValue(sheet, ref)
		headers = append(headers, val)
	}

	// Build type map by inspecting the first data row (or all headers if no data).
	colTypes := make(map[string]string, len(headers))
	for i, h := range headers {
		col := startCol + i
		typeName := "string"
		if startRow+1 <= endRow {
			ref, _ := excelize.CoordinatesToCellName(col, startRow+1)
			typeName = detectCellType(f, sheet, ref)
		}
		colTypes[h] = typeName
	}

	// Read data rows (everything after the header row).
	var dataRows []map[string]any
	dataStart := startRow + 1
	if opts.Limit > 0 && dataStart+opts.Limit-1 < endRow {
		endRow = dataStart + opts.Limit - 1
	}
	for r := dataStart; r <= endRow; r++ {
		row := make(map[string]any, len(headers))
		for i, h := range headers {
			col := startCol + i
			ref, _ := excelize.CoordinatesToCellName(col, r)
			row[h] = getTypedValue(f, sheet, ref)
		}
		dataRows = append(dataRows, row)
	}

	return &TypedReadResult{
		Sheet:   sheet,
		Range:   opts.Range,
		Headers: headers,
		Rows:    dataRows,
		Types:   colTypes,
	}, nil
}

// detectCellType returns the human-readable type of a cell, handling the
// CellTypeUnset → "number" ambiguity (in OOXML, cells without an explicit type
// attribute are numeric by default).
func detectCellType(f *excelize.File, sheet, ref string) string {
	cellType, _ := f.GetCellType(sheet, ref)
	switch cellType {
	case excelize.CellTypeBool:
		return "bool"
	case excelize.CellTypeDate:
		return "date"
	case excelize.CellTypeError:
		return "error"
	case excelize.CellTypeFormula:
		return "formula"
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		// CellTypeUnset means no explicit type in XML — default is numeric in OOXML.
		// Verify by attempting to parse the value; if it's not numeric, call it string.
		raw, _ := f.GetCellValue(sheet, ref)
		if raw == "" {
			return "string"
		}
		if _, err := strconv.ParseFloat(raw, 64); err == nil {
			return "number"
		}
		return "string"
	case excelize.CellTypeInlineString, excelize.CellTypeSharedString:
		return "string"
	default:
		return "string"
	}
}

// getTypedValue returns the cell value converted to its natural Go type.
// Numbers become float64 (or int when possible), bools become bool, everything else is string.
func getTypedValue(f *excelize.File, sheet, ref string) any {
	cellType, _ := f.GetCellType(sheet, ref)
	raw, _ := f.GetCellValue(sheet, ref)

	switch cellType {
	case excelize.CellTypeBool:
		return strings.ToLower(raw) == "true"
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		if raw == "" {
			return raw
		}
		if v, err := strconv.ParseFloat(raw, 64); err == nil {
			if v == float64(int(v)) && !strings.Contains(raw, ".") {
				return int(v)
			}
			return v
		}
		return raw
	default:
		return raw
	}
}

// readRange returns the rectangular block bounded by the A1-style range.
func readRange(f *excelize.File, sheet, rangeRef string) ([][]string, error) {
	parts := strings.Split(rangeRef, ":")
	if len(parts) != 2 {
		return nil, fmt.Errorf("invalid range: %s (expected e.g. A1:D10)", rangeRef)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[0]))
	if err != nil {
		return nil, fmt.Errorf("invalid start cell: %w", err)
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[1]))
	if err != nil {
		return nil, fmt.Errorf("invalid end cell: %w", err)
	}
	if startCol > endCol || startRow > endRow {
		return nil, fmt.Errorf("invalid range: start cell must be top-left of end cell")
	}

	out := make([][]string, 0, endRow-startRow+1)
	for r := startRow; r <= endRow; r++ {
		row := make([]string, 0, endCol-startCol+1)
		for c := startCol; c <= endCol; c++ {
			ref, _ := excelize.CoordinatesToCellName(c, r)
			val, _ := f.GetCellValue(sheet, ref)
			row = append(row, val)
		}
		out = append(out, row)
	}
	return out, nil
}

// CellWrite describes a single cell mutation for WriteCells.
type CellWrite struct {
	Sheet   string `json:"sheet"`
	Ref     string `json:"ref"`
	Value   any    `json:"value"`
	Formula string `json:"formula,omitempty"`
}

// WriteCells applies a list of cell mutations, then saves the workbook in place.
func WriteCells(path string, cells []CellWrite) error {
	if len(cells) == 0 {
		return errors.New("no cells to write")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	for _, c := range cells {
		sheet := c.Sheet
		if sheet == "" {
			sheet = f.GetSheetName(0)
		}
		if !sheetExists(f, sheet) {
			if _, err := f.NewSheet(sheet); err != nil {
				return fmt.Errorf("creating sheet %s: %w", sheet, err)
			}
		}
		if c.Formula != "" {
			if err := f.SetCellFormula(sheet, c.Ref, c.Formula); err != nil {
				return fmt.Errorf("setting formula at %s!%s: %w", sheet, c.Ref, err)
			}
			continue
		}
		if err := f.SetCellValue(sheet, c.Ref, c.Value); err != nil {
			return fmt.Errorf("setting value at %s!%s: %w", sheet, c.Ref, err)
		}
	}
	return f.Save()
}

// AppendRows appends rows to the end of the named sheet. The sheet is created if missing.
// Returns the row number where the first new row landed.
func AppendRows(path, sheet string, rows [][]any) (int, error) {
	if len(rows) == 0 {
		return 0, errors.New("no rows to append")
	}
	f, err := open(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		if _, err := f.NewSheet(sheet); err != nil {
			return 0, err
		}
	}

	existing, _ := f.GetRows(sheet)
	startRow := len(existing) + 1
	for r, row := range rows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, startRow+r)
			if err := f.SetCellValue(sheet, ref, val); err != nil {
				return 0, err
			}
		}
	}
	return startRow, f.Save()
}

// Search scans the entire workbook (or a specific sheet) for keyword. Case-insensitive.
func Search(path, keyword, onlySheet string, limit int) ([]SheetMatch, error) {
	if keyword == "" {
		return nil, errors.New("keyword cannot be empty")
	}
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	needle := strings.ToLower(keyword)
	var matches []SheetMatch

	sheets := f.GetSheetList()
	if onlySheet != "" {
		if !sheetExists(f, onlySheet) {
			return nil, fmt.Errorf("sheet not found: %s", onlySheet)
		}
		sheets = []string{onlySheet}
	}

	for _, sheet := range sheets {
		rows, err := f.GetRows(sheet)
		if err != nil {
			continue
		}
		for r, row := range rows {
			for c, cell := range row {
				if strings.Contains(strings.ToLower(cell), needle) {
					ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
					formula, _ := f.GetCellFormula(sheet, ref)
					matches = append(matches, SheetMatch{
						Sheet: sheet, Ref: ref, Value: cell, Formula: formula,
					})
					if limit > 0 && len(matches) >= limit {
						return matches, nil
					}
				}
			}
		}
	}
	return matches, nil
}

// Create builds a new workbook from the given spec at path. Refuses to overwrite by default.
func Create(path string, spec CreateSpec, overwrite bool) error {
	if !overwrite {
		if _, err := os.Stat(path); err == nil {
			return fmt.Errorf("file already exists (use --force to overwrite): %s", path)
		}
	}

	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	defaultSheet := f.GetSheetName(0)
	defaultDeleted := false

	if len(spec.Sheets) == 0 {
		spec.Sheets = []SheetSpec{{Name: "Sheet1"}}
	}

	for i, s := range spec.Sheets {
		name := s.Name
		if name == "" {
			name = fmt.Sprintf("Sheet%d", i+1)
		}
		if i == 0 && name != defaultSheet {
			if _, err := f.NewSheet(name); err != nil {
				return err
			}
		} else if i > 0 {
			if _, err := f.NewSheet(name); err != nil {
				return err
			}
		}
		for r, row := range s.Rows {
			for c, val := range row {
				ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
				if err := f.SetCellValue(name, ref, val); err != nil {
					return err
				}
			}
		}
		for c, col := range s.Cols {
			if col.Width > 0 {
				colName, _ := excelize.ColumnNumberToName(c + 1)
				_ = f.SetColWidth(name, colName, colName, col.Width)
			}
		}
	}

	if !defaultDeleted && len(spec.Sheets) > 0 && spec.Sheets[0].Name != defaultSheet && sheetExists(f, defaultSheet) {
		_ = f.DeleteSheet(defaultSheet)
	}

	if err := f.SaveAs(path); err != nil {
		return err
	}
	return nil
}

// ToCSV exports one sheet to a CSV file using UTF-8.
//
// withBOM=true prepends a UTF-8 BOM (\uFEFF). Microsoft Excel on Windows uses
// the BOM as a heuristic to interpret a CSV as UTF-8; without it CJK and other
// non-ASCII characters are rendered as garbage. Most other CSV readers ignore
// the BOM safely, so withBOM=true is the recommended default for human-facing
// exports.
func ToCSV(path, sheet, dest string, withBOM bool) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return err
	}

	out, err := os.Create(dest)
	if err != nil {
		return err
	}
	defer func() { _ = out.Close() }()

	if withBOM {
		if _, err := out.Write([]byte{0xEF, 0xBB, 0xBF}); err != nil {
			return err
		}
	}

	w := csv.NewWriter(out)
	defer w.Flush()
	for _, row := range rows {
		if err := w.Write(row); err != nil {
			return err
		}
	}
	return nil
}

// FromCSV imports one CSV file into a freshly created xlsx workbook.
// Sheet name defaults to "Sheet1". The CSV is read as UTF-8; a leading BOM is stripped.
func FromCSV(csvPath, xlsxPath, sheetName string, overwrite bool) (int, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	if !overwrite {
		if _, err := os.Stat(xlsxPath); err == nil {
			return 0, fmt.Errorf("file already exists (use --force to overwrite): %s", xlsxPath)
		}
	}

	in, err := os.Open(csvPath)
	if err != nil {
		return 0, err
	}
	defer func() { _ = in.Close() }()

	r := csv.NewReader(in)
	r.FieldsPerRecord = -1

	first := true
	rows := [][]string{}
	for {
		row, err := r.Read()
		if err != nil {
			if errors.Is(err, io.EOF) {
				break
			}
			return 0, err
		}
		if first && len(row) > 0 {
			row[0] = strings.TrimPrefix(row[0], "\uFEFF")
			first = false
		}
		rows = append(rows, row)
	}

	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	defaultName := f.GetSheetName(0)
	if sheetName != defaultName {
		if _, err := f.NewSheet(sheetName); err != nil {
			return 0, err
		}
		if err := f.DeleteSheet(defaultName); err != nil {
			return 0, err
		}
	}

	for r, row := range rows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellValue(sheetName, ref, val); err != nil {
				return 0, err
			}
		}
	}

	if err := f.SaveAs(xlsxPath); err != nil {
		return 0, err
	}
	return len(rows), nil
}

// ReadCell returns the value (and formula, if any) of a single cell. Cheap path
// for AI Agents that only need one value (avoids loading every row).
func ReadCell(path, sheet, ref string) (string, string, error) {
	f, err := open(path)
	if err != nil {
		return "", "", err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return "", "", fmt.Errorf("sheet not found: %s", sheet)
	}
	val, err := f.GetCellValue(sheet, ref)
	if err != nil {
		return "", "", err
	}
	formula, _ := f.GetCellFormula(sheet, ref)
	return val, formula, nil
}

// RenameSheet renames `from` to `to`. Both are sheet names. The sheet order
// is preserved; the underlying GetSheetIndex/SetSheetName excelize APIs handle this.
func RenameSheet(path, from, to string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if !sheetExists(f, from) {
		return fmt.Errorf("sheet not found: %s", from)
	}
	if from == to {
		return nil
	}
	if sheetExists(f, to) {
		return fmt.Errorf("target sheet name already exists: %s", to)
	}
	if err := f.SetSheetName(from, to); err != nil {
		return err
	}
	return f.Save()
}

// DeleteSheet removes a worksheet. Excel requires at least one sheet, so the
// last remaining sheet cannot be deleted.
func DeleteSheet(path, sheet string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if len(f.GetSheetList()) <= 1 {
		return errors.New("cannot delete the last remaining sheet")
	}
	if err := f.DeleteSheet(sheet); err != nil {
		return err
	}
	return f.Save()
}

// CopySheet duplicates `from` into a new sheet named `to`. Cell styles and
// formatting are copied; cell values are NOT copied (excelize limitation).
// Returns the newly created sheet's index.
func CopySheet(path, from, to string) (int, error) {
	f, err := open(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = f.Close() }()

	if !sheetExists(f, from) {
		return 0, fmt.Errorf("source sheet not found: %s", from)
	}
	if sheetExists(f, to) {
		return 0, fmt.Errorf("target sheet already exists: %s", to)
	}
	idx, err := f.NewSheet(to)
	if err != nil {
		return 0, err
	}
	srcIdx, err := f.GetSheetIndex(from)
	if err != nil {
		return 0, err
	}
	if err := f.CopySheet(srcIdx, idx); err != nil {
		return 0, err
	}
	return idx, f.Save()
}

// sheetExists is the canonical "is this sheet name present" check.
func sheetExists(f *excelize.File, name string) bool {
	for _, n := range f.GetSheetList() {
		if n == name {
			return true
		}
	}
	return false
}

// splitRange splits "A1:C3" into ("A1", "C3"). If there is no colon,
// both return values are the original string (single-cell range).
func splitRange(r string) (string, string) {
	parts := strings.SplitN(r, ":", 2)
	if len(parts) == 2 {
		return parts[0], parts[1]
	}
	return r, r
}

// MapSheetInfo converts our domain types into output.FlatSheet for JSON consistency.
func MapSheetInfo(s SheetInfo) output.FlatSheet {
	return output.FlatSheet{
		Name:      s.Name,
		Index:     s.Index,
		Rows:      s.Rows,
		Cols:      s.Cols,
		Hidden:    s.Hidden,
		Dimension: s.Dimension,
	}
}

// ---------------------------------------------------------------------------
// Write-path engine functions (style, structure, charts, formatting)
// ---------------------------------------------------------------------------

// StyleSpec describes the visual properties to apply to a cell range.
type StyleSpec struct {
	Bold         bool         `json:"bold,omitempty"`
	Italic       bool         `json:"italic,omitempty"`
	Underline    bool         `json:"underline,omitempty"`
	Strike       bool         `json:"strike,omitempty"`   // strikethrough
	FontSize     float64      `json:"fontSize,omitempty"` // points; 0 = default
	FontFamily   string       `json:"fontFamily,omitempty"`
	BgColor      string       `json:"bgColor,omitempty"`   // hex e.g. "FFFF00"
	TextColor    string       `json:"textColor,omitempty"` // hex e.g. "FF0000"
	NumberFormat string       `json:"numberFormat,omitempty"`
	Align        string       `json:"align,omitempty"`  // left, center, right
	Valign       string       `json:"valign,omitempty"` // top, center, bottom
	WrapText     bool         `json:"wrapText,omitempty"`
	TextRotation int          `json:"textRotation,omitempty"` // 0-90 degrees, or 255 for vertical text
	Indent       int          `json:"indent,omitempty"`       // indent level (each level ≈ 3 spaces)
	ShrinkToFit  bool         `json:"shrinkToFit,omitempty"`  // shrink font to fit cell width
	Borders      []BorderSpec `json:"borders,omitempty"`
}

// BorderSpec is one border edge.
type BorderSpec struct {
	Type  string `json:"type"`  // left, right, top, bottom, diagonal
	Color string `json:"color"` // hex
	Style int    `json:"style"` // excelize border style (1=thin, 2=medium, etc.)
}

// StyleCells applies formatting to a range of cells.
func StyleCells(path, sheet, cellRange string, spec StyleSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	style := excelize.Style{}
	font := &excelize.Font{}
	hasFont := false
	if spec.Bold {
		font.Bold = true
		hasFont = true
	}
	if spec.Italic {
		font.Italic = true
		hasFont = true
	}
	if spec.Underline {
		font.Underline = "single"
		hasFont = true
	}
	if spec.FontSize > 0 {
		font.Size = spec.FontSize
		hasFont = true
	}
	if spec.FontFamily != "" {
		font.Family = spec.FontFamily
		hasFont = true
	}
	if spec.TextColor != "" {
		font.Color = spec.TextColor
		hasFont = true
	}
	if spec.Strike {
		font.Strike = true
		hasFont = true
	}
	if hasFont {
		style.Font = font
	}

	if spec.BgColor != "" {
		style.Fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{spec.BgColor},
		}
	}

	if spec.NumberFormat != "" {
		style.NumFmt = 0 // will be overridden
		style.CustomNumFmt = &spec.NumberFormat
	}

	if spec.Align != "" || spec.Valign != "" || spec.WrapText || spec.TextRotation != 0 || spec.Indent > 0 || spec.ShrinkToFit {
		alignment := &excelize.Alignment{WrapText: spec.WrapText, ShrinkToFit: spec.ShrinkToFit, Indent: spec.Indent, TextRotation: spec.TextRotation}
		switch spec.Align {
		case "left":
			alignment.Horizontal = "left"
		case "center":
			alignment.Horizontal = "center"
		case "right":
			alignment.Horizontal = "right"
		}
		switch spec.Valign {
		case "top":
			alignment.Vertical = "top"
		case "center":
			alignment.Vertical = "center"
		case "bottom":
			alignment.Vertical = "bottom"
		}
		style.Alignment = alignment
	}

	if len(spec.Borders) > 0 {
		borders := []excelize.Border{}
		for _, b := range spec.Borders {
			borders = append(borders, excelize.Border{
				Type:  b.Type,
				Color: b.Color,
				Style: b.Style,
			})
		}
		style.Border = borders
	}

	styleID, err := f.NewStyle(&style)
	if err != nil {
		return fmt.Errorf("creating style: %w", err)
	}

	topLeft, bottomRight := splitRange(cellRange)
	if err := f.SetCellStyle(sheet, topLeft, bottomRight, styleID); err != nil {
		return fmt.Errorf("applying style to %s: %w", cellRange, err)
	}
	return f.Save()
}

// InsertRows inserts `count` blank rows after `afterRow` (1-based).
func InsertRows(path, sheet string, afterRow, count int) error {
	if count <= 0 {
		return errors.New("count must be > 0")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if err := f.InsertRows(sheet, afterRow+1, count); err != nil {
		return err
	}
	return f.Save()
}

// InsertCols inserts `count` blank columns after `afterCol` (1-based).
func InsertCols(path, sheet string, afterCol, count int) error {
	if count <= 0 {
		return errors.New("count must be > 0")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	colName, err := excelize.ColumnNumberToName(afterCol + 1)
	if err != nil {
		return fmt.Errorf("invalid column: %w", err)
	}
	if err := f.InsertCols(sheet, colName, count); err != nil {
		return err
	}
	return f.Save()
}

// DeleteRows deletes `count` rows starting at `fromRow` (1-based).
func DeleteRows(path, sheet string, fromRow, count int) error {
	if count <= 0 {
		return errors.New("count must be > 0")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	for i := 0; i < count; i++ {
		if err := f.RemoveRow(sheet, fromRow); err != nil {
			return err
		}
	}
	return f.Save()
}

// DeleteCols deletes `count` columns starting at `fromCol` (1-based).
func DeleteCols(path, sheet string, fromCol, count int) error {
	if count <= 0 {
		return errors.New("count must be > 0")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	for i := 0; i < count; i++ {
		colName, err := excelize.ColumnNumberToName(fromCol)
		if err != nil {
			return fmt.Errorf("invalid column: %w", err)
		}
		if err := f.RemoveCol(sheet, colName); err != nil {
			return err
		}
	}
	return f.Save()
}

// Sort reads a range, sorts rows by the specified column, and writes back.
// sortByCol is 1-based (1=first column of the range).
func Sort(path, sheet, cellRange string, sortByCol int, ascending bool) error {
	if sortByCol < 1 {
		return errors.New("sortByCol must be >= 1")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	parts := strings.Split(cellRange, ":")
	if len(parts) != 2 {
		return fmt.Errorf("invalid range: %s (expected e.g. A1:D10)", cellRange)
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[0]))
	if err != nil {
		return fmt.Errorf("invalid start cell: %w", err)
	}
	endCol, _, err := excelize.CellNameToCoordinates(strings.TrimSpace(parts[1]))
	if err != nil {
		return fmt.Errorf("invalid end cell: %w", err)
	}
	rangeCols := endCol - startCol + 1
	if sortByCol > rangeCols {
		return fmt.Errorf("sortByCol %d exceeds range width %d", sortByCol, rangeCols)
	}

	data, err := readRange(f, sheet, cellRange)
	if err != nil {
		return err
	}
	if len(data) <= 1 {
		return nil // nothing to sort
	}

	sortIdx := sortByCol - 1
	sort.Slice(data[1:], func(i, j int) bool {
		a, b := "", ""
		rowI, rowJ := data[1+i], data[1+j]
		if sortIdx < len(rowI) {
			a = rowI[sortIdx]
		}
		if sortIdx < len(rowJ) {
			b = rowJ[sortIdx]
		}
		aNum, aIsNum := parseNumber(a)
		bNum, bIsNum := parseNumber(b)
		if aIsNum && bIsNum {
			if ascending {
				return aNum < bNum
			}
			return aNum > bNum
		}
		if ascending {
			return strings.ToLower(a) < strings.ToLower(b)
		}
		return strings.ToLower(a) > strings.ToLower(b)
	})

	for r, row := range data {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(startCol+c, startRow+r)
			_ = f.SetCellValue(sheet, ref, val)
		}
	}
	return f.Save()
}

// parseNumber attempts to parse s as a float64. Returns (value, true) on success.
func parseNumber(s string) (float64, bool) {
	if s == "" {
		return 0, false
	}
	// Try int first (common case)
	if v, err := strconv.Atoi(s); err == nil {
		return float64(v), true
	}
	// Try float
	if v, err := strconv.ParseFloat(s, 64); err == nil {
		return v, true
	}
	return 0, false
}

// Freeze sets freeze panes at the given cell (e.g. "A2" freezes the top row).
func Freeze(path, sheet, cell string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	panes := &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      0,
		TopLeftCell: cell,
		ActivePane:  "bottomLeft",
	}
	// Parse cell to get X/Y split
	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("invalid cell: %w", err)
	}
	panes.XSplit = col - 1
	panes.YSplit = row - 1
	if panes.XSplit > 0 && panes.YSplit > 0 {
		panes.ActivePane = "bottomRight"
	} else if panes.XSplit > 0 {
		panes.ActivePane = "topRight"
	}

	if err := f.SetPanes(sheet, panes); err != nil {
		return err
	}
	return f.Save()
}

// MergeCells merges a rectangular range.
func MergeCells(path, sheet, cellRange string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	topLeft, bottomRight := splitRange(cellRange)
	if err := f.MergeCell(sheet, topLeft, bottomRight); err != nil {
		return err
	}
	return f.Save()
}

// UnmergeCells unmerges a previously merged range.
func UnmergeCells(path, sheet, cellRange string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	topLeft, bottomRight := splitRange(cellRange)
	if err := f.UnmergeCell(sheet, topLeft, bottomRight); err != nil {
		return err
	}
	return f.Save()
}

// SetColWidth sets the width of a single column (1-based).
func SetColWidth(path, sheet string, col int, width float64) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	colName, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return fmt.Errorf("invalid column: %w", err)
	}
	if err := f.SetColWidth(sheet, colName, colName, width); err != nil {
		return err
	}
	return f.Save()
}

// SetRowHeight sets the height of a single row (1-based).
func SetRowHeight(path, sheet string, row int, height float64) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if err := f.SetRowHeight(sheet, row, height); err != nil {
		return err
	}
	return f.Save()
}

// ChartSpec describes a chart to create.
type ChartSpec struct {
	Type      string `json:"type"`      // col, bar, line, area, pie, scatter
	Cell      string `json:"cell"`      // top-left anchor cell
	DataRange string `json:"dataRange"` // e.g. "A1:B5"
	Title     string `json:"title,omitempty"`
	Width     int    `json:"width,omitempty"`  // default 12
	Height    int    `json:"height,omitempty"` // default 8
}

// AddChart creates a chart on the specified sheet.
func AddChart(path, sheet string, spec ChartSpec) error {
	if spec.DataRange == "" {
		return errors.New("dataRange is required")
	}
	if spec.Cell == "" {
		return errors.New("cell (anchor) is required")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	width := uint(spec.Width)
	if width <= 0 {
		width = 12
	}
	height := uint(spec.Height)
	if height <= 0 {
		height = 8
	}

	parts := strings.Split(spec.DataRange, ":")
	if len(parts) != 2 {
		return fmt.Errorf("invalid dataRange: %s (expected e.g. A1:B5)", spec.DataRange)
	}

	var chartType excelize.ChartType
	switch strings.ToLower(spec.Type) {
	case "col", "column":
		chartType = excelize.Col
	case "bar":
		chartType = excelize.Bar
	case "line":
		chartType = excelize.Line
	case "area":
		chartType = excelize.Area
	case "pie":
		chartType = excelize.Pie
	case "scatter":
		chartType = excelize.Scatter
	default:
		return fmt.Errorf("unsupported chart type: %s (use col, bar, line, area, pie, scatter)", spec.Type)
	}

	chart := &excelize.Chart{
		Type: chartType,
		Series: []excelize.ChartSeries{
			{Categories: fmt.Sprintf("'%s'!%s", sheet, spec.DataRange),
				Values: fmt.Sprintf("'%s'!%s", sheet, spec.DataRange)},
		},
		Dimension: excelize.ChartDimension{
			Width:  width,
			Height: height,
		},
	}
	if spec.Title != "" {
		chart.Title = []excelize.RichTextRun{
			{Text: spec.Title},
		}
	}

	if err := f.AddChart(sheet, spec.Cell, chart); err != nil {
		return fmt.Errorf("adding chart: %w", err)
	}
	return f.Save()
}

// AddImage inserts an image anchored at a cell.
func AddImage(path, sheet, cell, imagePath string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if err := f.AddPicture(sheet, cell, imagePath, nil); err != nil {
		return err
	}
	return f.Save()
}

// CondSpec describes a conditional formatting rule.
type CondSpec struct {
	Type      string `json:"type"` // cell, top, bottom, average, duplicate, unique, formula
	MinVal    string `json:"minVal,omitempty"`
	MaxVal    string `json:"maxVal,omitempty"`
	MinColor  string `json:"minColor,omitempty"` // hex
	MaxColor  string `json:"maxColor,omitempty"` // hex
	BgColor   string `json:"bgColor,omitempty"`
	TextColor string `json:"textColor,omitempty"`
	Criteria  string `json:"criteria,omitempty"` // >, <, >=, <=, ==, !=
	Value     string `json:"value,omitempty"`
}

// ConditionalFormat applies conditional formatting to a range.
func ConditionalFormat(path, sheet, cellRange string, spec CondSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	// Build a style for the conditional format
	style := excelize.Style{}
	if spec.BgColor != "" {
		style.Fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{spec.BgColor},
		}
	}
	if spec.TextColor != "" {
		style.Font = &excelize.Font{Color: spec.TextColor}
	}
	styleID, err := f.NewStyle(&style)
	if err != nil {
		return fmt.Errorf("creating conditional format style: %w", err)
	}

	format := excelize.ConditionalFormatOptions{
		Format: styleID,
	}
	switch strings.ToLower(spec.Type) {
	case "cell":
		format.Type = "cell"
		format.Criteria = spec.Criteria
		if spec.Value != "" {
			format.Value = spec.Value
		}
	case "top":
		format.Type = "top10"
		format.Criteria = ">="
		if spec.Value != "" {
			format.Value = spec.Value
		}
	case "bottom":
		format.Type = "top10"
		format.Percent = true
		format.Criteria = "<="
		if spec.Value != "" {
			format.Value = spec.Value
		}
	case "average":
		format.Type = "aboveAverage"
	case "duplicate":
		format.Type = "duplicateValues"
	case "unique":
		format.Type = "uniqueValues"
	case "formula":
		format.Type = "formula"
		format.Value = spec.Value
	default:
		return fmt.Errorf("unsupported conditional format type: %s", spec.Type)
	}

	if err := f.SetConditionalFormat(sheet, cellRange, []excelize.ConditionalFormatOptions{format}); err != nil {
		return err
	}
	return f.Save()
}

// SetAutoFilter applies auto-filter to a range.
func SetAutoFilter(path, sheet, cellRange string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if err := f.AutoFilter(sheet, cellRange, []excelize.AutoFilterOptions{}); err != nil {
		return err
	}
	return f.Save()
}

// SetSheetVisible controls whether a worksheet is visible.
func SetSheetVisible(path, sheet string, visible bool) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if err := f.SetSheetVisible(sheet, visible); err != nil {
		return err
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Feature 2: Formula generation
// ---------------------------------------------------------------------------

// FormulaSpec describes a formula to generate.
type FormulaSpec struct {
	Column     string `json:"column"`               // e.g. "B" or column letter
	Func       string `json:"func"`                 // SUM, AVERAGE, COUNT, COUNTA, MIN, MAX
	OutputCell string `json:"outputCell,omitempty"` // e.g. "B101" (auto-detected if empty)
	DataRange  string `json:"dataRange,omitempty"`  // e.g. "B2:B100" (auto-detected if empty)
	Criteria   string `json:"criteria,omitempty"`   // for COUNTIF/SUMIF
}

// FormulaResult is the output of GenerateFormula.
type FormulaResult struct {
	Cell      string `json:"cell"`
	Formula   string `json:"formula"`
	DataRows  int    `json:"dataRows"`
	DataRange string `json:"dataRange"`
}

// GenerateFormula creates a formula for a column. It auto-detects the data range
// if not provided. Supported funcs: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX,
// COUNTIF, SUMIF.
func GenerateFormula(path, sheet string, spec FormulaSpec) (*FormulaResult, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return nil, fmt.Errorf("sheet not found: %s", sheet)
	}

	col := spec.Column
	if col == "" {
		return nil, errors.New("column is required")
	}
	// Normalize: allow "2" as column number.
	if n, err := strconv.Atoi(col); err == nil {
		colName, err := excelize.ColumnNumberToName(n)
		if err != nil {
			return nil, fmt.Errorf("invalid column number: %d", n)
		}
		col = colName
	}
	col = strings.ToUpper(col)

	dataRange := spec.DataRange
	dataRows := 0

	if dataRange == "" {
		// Auto-detect: find last non-empty row in the column.
		rows, _ := f.GetRows(sheet)
		colIdx := 0
		for i := 0; i < len(col); i++ {
			colIdx = colIdx*26 + int(col[i]-'A') + 1
		}
		lastRow := 0
		for r, row := range rows {
			if colIdx-1 < len(row) && row[colIdx-1] != "" {
				lastRow = r + 1 // 1-based
			}
		}
		if lastRow < 2 {
			lastRow = 2 // minimum: row 2 (after header)
		}
		dataRange = fmt.Sprintf("%s2:%s%d", col, col, lastRow)
		dataRows = lastRow - 1
	} else {
		// Parse the provided range to count rows.
		parts := strings.Split(dataRange, ":")
		if len(parts) == 2 {
			_, startRow, _ := excelize.CellNameToCoordinates(parts[0])
			_, endRow, _ := excelize.CellNameToCoordinates(parts[1])
			dataRows = endRow - startRow + 1
		}
	}

	outputCell := spec.OutputCell
	if outputCell == "" {
		// Place formula one row below the data range.
		parts := strings.Split(dataRange, ":")
		if len(parts) == 2 {
			_, endRow, _ := excelize.CellNameToCoordinates(parts[1])
			outputCell = fmt.Sprintf("%s%d", col, endRow+1)
		} else {
			outputCell = col + "2"
		}
	}

	// Build the formula.
	fn := strings.ToUpper(spec.Func)
	var formula string
	switch fn {
	case "SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX":
		formula = fmt.Sprintf("=%s(%s)", fn, dataRange)
	case "COUNTIF":
		if spec.Criteria == "" {
			return nil, errors.New("criteria is required for COUNTIF")
		}
		formula = fmt.Sprintf("=COUNTIF(%s,%q)", dataRange, spec.Criteria)
	case "SUMIF":
		if spec.Criteria == "" {
			return nil, errors.New("criteria is required for SUMIF")
		}
		formula = fmt.Sprintf("=SUMIF(%s,%q)", dataRange, spec.Criteria)
	default:
		return nil, fmt.Errorf("unsupported function: %s (use SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, COUNTIF, SUMIF)", spec.Func)
	}

	return &FormulaResult{
		Cell:      outputCell,
		Formula:   formula,
		DataRows:  dataRows,
		DataRange: dataRange,
	}, nil
}

// GenerateFormulas generates multiple formulas in batch.
func GenerateFormulas(path, sheet string, specs []FormulaSpec) ([]FormulaResult, error) {
	results := make([]FormulaResult, 0, len(specs))
	for _, spec := range specs {
		r, err := GenerateFormula(path, sheet, spec)
		if err != nil {
			return results, err
		}
		results = append(results, *r)
	}
	return results, nil
}

// ApplyFormulas writes generated formulas to the workbook.
func ApplyFormulas(path, sheet string, results []FormulaResult) error {
	cells := make([]CellWrite, 0, len(results))
	for _, r := range results {
		cells = append(cells, CellWrite{
			Sheet:   sheet,
			Ref:     r.Cell,
			Formula: r.Formula,
		})
	}
	return WriteCells(path, cells)
}

// ---------------------------------------------------------------------------
// Feature 3: Workbook copy
// ---------------------------------------------------------------------------

// CopyWorkbook deep-copies a workbook (all sheets, data, styles, charts, images).
func CopyWorkbook(src, dst string, overwrite bool) error {
	if !overwrite {
		if _, err := os.Stat(dst); err == nil {
			return fmt.Errorf("file already exists (use --force to overwrite): %s", dst)
		}
	}
	f, err := excelize.OpenFile(src)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	return f.SaveAs(dst)
}

// ---------------------------------------------------------------------------
// Feature 4: Data validation
// ---------------------------------------------------------------------------

// ValidationSpec describes a data validation rule.
type ValidationSpec struct {
	Type        string   `json:"type"`               // list, whole, decimal, date, time, textLength, custom
	Range       string   `json:"range"`              // e.g. "D2:D100"
	Operator    string   `json:"operator,omitempty"` // between, equal, greaterThan, etc.
	Min         string   `json:"min,omitempty"`      // for numeric/date ranges
	Max         string   `json:"max,omitempty"`      // for numeric/date ranges
	List        []string `json:"list,omitempty"`     // for dropdown list
	Formula     string   `json:"formula,omitempty"`  // for custom validation
	ErrorMsg    string   `json:"errorMsg,omitempty"`
	ErrorTitle  string   `json:"errorTitle,omitempty"`
	PromptMsg   string   `json:"promptMsg,omitempty"`
	PromptTitle string   `json:"promptTitle,omitempty"`
	AllowBlank  bool     `json:"allowBlank,omitempty"`
}

// AddValidation adds a data validation rule to a worksheet.
func AddValidation(path, sheet string, spec ValidationSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	if spec.Range == "" {
		return errors.New("range is required")
	}

	dv := excelize.NewDataValidation(spec.AllowBlank)
	dv.Sqref = spec.Range

	if spec.Type == "list" {
		if len(spec.List) == 0 {
			return errors.New("list is required for type 'list'")
		}
		if err := dv.SetDropList(spec.List); err != nil {
			return fmt.Errorf("setting drop list: %w", err)
		}
	} else if spec.Type == "custom" {
		if spec.Formula == "" {
			return errors.New("formula is required for type 'custom'")
		}
		dv.Type = "custom"
		dv.Formula1 = spec.Formula
	} else {
		vType, err := parseValidationType(spec.Type)
		if err != nil {
			return err
		}
		op, err := parseValidationOperator(spec.Operator)
		if err != nil {
			return err
		}
		if err := dv.SetRange(spec.Min, spec.Max, vType, op); err != nil {
			return fmt.Errorf("setting range: %w", err)
		}
	}

	if spec.ErrorMsg != "" {
		dv.SetError(excelize.DataValidationErrorStyleStop, spec.ErrorTitle, spec.ErrorMsg)
	}
	if spec.PromptMsg != "" {
		dv.SetInput(spec.PromptTitle, spec.PromptMsg)
	}

	if err := f.AddDataValidation(sheet, dv); err != nil {
		return fmt.Errorf("adding data validation: %w", err)
	}
	return f.Save()
}

func parseValidationType(s string) (excelize.DataValidationType, error) {
	switch strings.ToLower(s) {
	case "whole":
		return excelize.DataValidationTypeWhole, nil
	case "decimal":
		return excelize.DataValidationTypeDecimal, nil
	case "date":
		return excelize.DataValidationTypeDate, nil
	case "time":
		return excelize.DataValidationTypeTime, nil
	case "textlength", "text-length":
		return excelize.DataValidationTypeTextLength, nil
	default:
		return 0, fmt.Errorf("unsupported validation type: %s (use whole, decimal, date, time, textLength)", s)
	}
}

func parseValidationOperator(s string) (excelize.DataValidationOperator, error) {
	if s == "" {
		return excelize.DataValidationOperatorBetween, nil
	}
	switch strings.ToLower(s) {
	case "between", "":
		return excelize.DataValidationOperatorBetween, nil
	case "equal", "eq":
		return excelize.DataValidationOperatorEqual, nil
	case "greaterthan", "greater-than", "gt":
		return excelize.DataValidationOperatorGreaterThan, nil
	case "greaterthanorequal", "gte", "ge":
		return excelize.DataValidationOperatorGreaterThanOrEqual, nil
	case "lessthan", "less-than", "lt":
		return excelize.DataValidationOperatorLessThan, nil
	case "lessthanorequal", "lte", "le":
		return excelize.DataValidationOperatorLessThanOrEqual, nil
	case "notbetween", "not-between":
		return excelize.DataValidationOperatorNotBetween, nil
	case "notequal", "not-equal", "ne":
		return excelize.DataValidationOperatorNotEqual, nil
	default:
		return 0, fmt.Errorf("unsupported validation operator: %s", s)
	}
}

// ---------------------------------------------------------------------------
// Feature 5: Batch style
// ---------------------------------------------------------------------------

// BatchStyleEntry pairs a cell range with its style spec.
type BatchStyleEntry struct {
	Range     string    `json:"range"`
	StyleSpec StyleSpec `json:"style"`
}

// StyleCellsBatch applies multiple styles in a single open/save cycle.
func StyleCellsBatch(path, sheet string, entries []BatchStyleEntry) error {
	if len(entries) == 0 {
		return errors.New("no style entries provided")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	for _, entry := range entries {
		if entry.Range == "" {
			return errors.New("range is required for each style entry")
		}
		spec := entry.StyleSpec
		style := excelize.Style{}
		font := &excelize.Font{}
		hasFont := false
		if spec.Bold {
			font.Bold = true
			hasFont = true
		}
		if spec.Italic {
			font.Italic = true
			hasFont = true
		}
		if spec.Underline {
			font.Underline = "single"
			hasFont = true
		}
		if spec.FontSize > 0 {
			font.Size = spec.FontSize
			hasFont = true
		}
		if spec.FontFamily != "" {
			font.Family = spec.FontFamily
			hasFont = true
		}
		if spec.TextColor != "" {
			font.Color = spec.TextColor
			hasFont = true
		}
		if hasFont {
			style.Font = font
		}
		if spec.BgColor != "" {
			style.Fill = excelize.Fill{
				Type:    "pattern",
				Pattern: 1,
				Color:   []string{spec.BgColor},
			}
		}
		if spec.NumberFormat != "" {
			style.CustomNumFmt = &spec.NumberFormat
		}
		if spec.Align != "" || spec.Valign != "" || spec.WrapText {
			alignment := &excelize.Alignment{WrapText: spec.WrapText}
			switch spec.Align {
			case "left":
				alignment.Horizontal = "left"
			case "center":
				alignment.Horizontal = "center"
			case "right":
				alignment.Horizontal = "right"
			}
			switch spec.Valign {
			case "top":
				alignment.Vertical = "top"
			case "center":
				alignment.Vertical = "center"
			case "bottom":
				alignment.Vertical = "bottom"
			}
			style.Alignment = alignment
		}
		if len(spec.Borders) > 0 {
			borders := []excelize.Border{}
			for _, b := range spec.Borders {
				borders = append(borders, excelize.Border{
					Type:  b.Type,
					Color: b.Color,
					Style: b.Style,
				})
			}
			style.Border = borders
		}
		styleID, err := f.NewStyle(&style)
		if err != nil {
			return fmt.Errorf("creating style for range %s: %w", entry.Range, err)
		}
		topLeft, bottomRight := splitRange(entry.Range)
		if err := f.SetCellStyle(sheet, topLeft, bottomRight, styleID); err != nil {
			return fmt.Errorf("applying style to %s: %w", entry.Range, err)
		}
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Feature 2b: Column-level formula generation (VLOOKUP, IF, cross-sheet, etc.)
// ---------------------------------------------------------------------------

// ColumnFormulaSpec describes a formula to apply down an entire column.
type ColumnFormulaSpec struct {
	Column   string `json:"column"`             // target column, e.g. "D"
	Func     string `json:"func"`               // VLOOKUP, IF, SUMIF, AVERAGEIF, CONCAT, CUSTOM
	StartRow int    `json:"startRow,omitempty"` // first data row (default: 2)
	EndRow   int    `json:"endRow,omitempty"`   // last data row (auto-detected if 0)

	// VLOOKUP parameters
	LookupCol  string `json:"lookupCol,omitempty"`  // column to search in, e.g. "A"
	ReturnCol  string `json:"returnCol,omitempty"`  // column to return from, e.g. "C"
	TableRange string `json:"tableRange,omitempty"` // e.g. "A2:C100" or "Sheet2!A:C"
	ExactMatch *bool  `json:"exactMatch,omitempty"` // default: true (0 = exact)

	// IF parameters
	Condition    string `json:"condition,omitempty"`   // e.g. "B{row}>100"
	ValueIfTrue  string `json:"valueIfTrue,omitempty"` // e.g. "\"Pass\""
	ValueIfFalse string `json:"valueIfFalse,omitempty"`

	// SUMIF / AVERAGEIF parameters
	CriteriaRange string `json:"criteriaRange,omitempty"` // e.g. "A2:A100"
	Criteria      string `json:"criteria,omitempty"`      // e.g. "\"Alice\""
	SumRange      string `json:"sumRange,omitempty"`      // e.g. "B2:B100"

	// CONCAT parameters
	ConcatCols []string `json:"concatCols,omitempty"` // e.g. ["A", "B"]
	Separator  string   `json:"separator,omitempty"`  // e.g. " "

	// Custom formula (the {row} placeholder is replaced per row)
	CustomFormula string `json:"customFormula,omitempty"` // e.g. "=B{row}*C{row}"

	// Cross-sheet reference
	SheetRef string `json:"sheetRef,omitempty"` // sheet for data range, e.g. "Master"
}

// ColumnFormulaResult is the output of GenerateColumnFormula.
type ColumnFormulaResult struct {
	Column    string `json:"column"`
	RowCount  int    `json:"rowCount"`
	SampleRow int    `json:"sampleRow"` // first data row
	Formula   string `json:"formula"`   // formula for sampleRow
}

// GenerateColumnFormula generates a formula for the first data row of a column.
// Returns the pattern; caller uses ApplyColumnFormulas to write all rows.
func GenerateColumnFormula(path, sheet string, spec ColumnFormulaSpec) (*ColumnFormulaResult, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return nil, fmt.Errorf("sheet not found: %s", sheet)
	}

	col := strings.ToUpper(spec.Column)
	if col == "" {
		return nil, errors.New("column is required")
	}
	if n, err := strconv.Atoi(col); err == nil {
		colName, err := excelize.ColumnNumberToName(n)
		if err != nil {
			return nil, fmt.Errorf("invalid column number: %d", n)
		}
		col = colName
	}

	startRow := spec.StartRow
	if startRow < 2 {
		startRow = 2
	}
	endRow := spec.EndRow
	if endRow == 0 {
		endRow = detectLastRow(f, sheet)
	}
	if endRow < startRow {
		endRow = startRow
	}
	rowCount := endRow - startRow + 1

	sampleFormula, err := buildRowFormula(spec, startRow)
	if err != nil {
		return nil, err
	}

	return &ColumnFormulaResult{
		Column:    col,
		RowCount:  rowCount,
		SampleRow: startRow,
		Formula:   sampleFormula,
	}, nil
}

// ApplyColumnFormulas writes formulas to every row in the range.
func ApplyColumnFormulas(path, sheet string, spec ColumnFormulaSpec) (int, error) {
	f, err := open(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}

	col := strings.ToUpper(spec.Column)
	if n, err := strconv.Atoi(col); err == nil {
		colName, _ := excelize.ColumnNumberToName(n)
		col = colName
	}

	startRow := spec.StartRow
	if startRow < 2 {
		startRow = 2
	}
	endRow := spec.EndRow
	if endRow == 0 {
		endRow = detectLastRow(f, sheet)
	}
	if endRow < startRow {
		endRow = startRow
	}

	cells := make([]CellWrite, 0, endRow-startRow+1)
	for r := startRow; r <= endRow; r++ {
		formula, err := buildRowFormula(spec, r)
		if err != nil {
			return 0, err
		}
		ref := fmt.Sprintf("%s%d", col, r)
		cells = append(cells, CellWrite{Sheet: sheet, Ref: ref, Formula: formula})
	}

	if err := WriteCells(path, cells); err != nil {
		return 0, err
	}
	return len(cells), nil
}

// buildRowFormula constructs a single formula for a given row number.
func buildRowFormula(spec ColumnFormulaSpec, row int) (string, error) {
	fn := strings.ToUpper(spec.Func)

	qualify := func(r string) string {
		if spec.SheetRef != "" && !strings.Contains(r, "!") {
			return fmt.Sprintf("'%s'!%s", spec.SheetRef, r)
		}
		return r
	}

	switch fn {
	case "VLOOKUP":
		if spec.LookupCol == "" || spec.ReturnCol == "" || spec.TableRange == "" {
			return "", errors.New("VLOOKUP requires lookupCol, returnCol, and tableRange")
		}
		lc := strings.ToUpper(spec.LookupCol)
		tableR := qualify(spec.TableRange)
		// Calculate column index within table range.
		parts := strings.Split(spec.TableRange, "!")
		rangePart := parts[len(parts)-1]
		rangeParts := strings.Split(rangePart, ":")
		colIdx := 1
		if len(rangeParts) == 2 {
			startCol := strings.TrimRight(rangeParts[0], "0123456789")
			rc := strings.ToUpper(spec.ReturnCol)
			sc := strings.ToUpper(startCol)
			colIdx = colNameToNumber(rc) - colNameToNumber(sc) + 1
		}
		matchType := "0"
		if spec.ExactMatch != nil && !*spec.ExactMatch {
			matchType = "1"
		}
		return fmt.Sprintf("=VLOOKUP(%s%d,%s,%d,%s)", lc, row, tableR, colIdx, matchType), nil

	case "IF":
		if spec.Condition == "" {
			return "", errors.New("IF requires condition")
		}
		cond := replaceRowPlaceholder(spec.Condition, row)
		vt := replaceRowPlaceholder(defaultStr(spec.ValueIfTrue, "\"\""), row)
		vf := replaceRowPlaceholder(defaultStr(spec.ValueIfFalse, "\"\""), row)
		return fmt.Sprintf("=IF(%s,%s,%s)", cond, vt, vf), nil

	case "SUMIF":
		if spec.CriteriaRange == "" || spec.Criteria == "" {
			return "", errors.New("SUMIF requires criteriaRange and criteria")
		}
		cr := qualify(replaceRowPlaceholder(spec.CriteriaRange, row))
		crit := replaceRowPlaceholder(spec.Criteria, row)
		if spec.SumRange != "" {
			sr := qualify(replaceRowPlaceholder(spec.SumRange, row))
			return fmt.Sprintf("=SUMIF(%s,%s,%s)", cr, crit, sr), nil
		}
		return fmt.Sprintf("=SUMIF(%s,%s)", cr, crit), nil

	case "AVERAGEIF":
		if spec.CriteriaRange == "" || spec.Criteria == "" {
			return "", errors.New("AVERAGEIF requires criteriaRange and criteria")
		}
		cr := qualify(replaceRowPlaceholder(spec.CriteriaRange, row))
		crit := replaceRowPlaceholder(spec.Criteria, row)
		if spec.SumRange != "" {
			sr := qualify(replaceRowPlaceholder(spec.SumRange, row))
			return fmt.Sprintf("=AVERAGEIF(%s,%s,%s)", cr, crit, sr), nil
		}
		return fmt.Sprintf("=AVERAGEIF(%s,%s)", cr, crit), nil

	case "CONCAT":
		if len(spec.ConcatCols) == 0 {
			return "", errors.New("CONCAT requires concatCols")
		}
		sep := spec.Separator
		parts := make([]string, len(spec.ConcatCols))
		for i, c := range spec.ConcatCols {
			parts[i] = fmt.Sprintf("%s%d", strings.ToUpper(c), row)
		}
		if sep == "" {
			return fmt.Sprintf("=CONCATENATE(%s)", strings.Join(parts, ",")), nil
		}
		var interleaved []string
		for i, p := range parts {
			if i > 0 {
				interleaved = append(interleaved, fmt.Sprintf("%q", sep))
			}
			interleaved = append(interleaved, p)
		}
		return fmt.Sprintf("=CONCATENATE(%s)", strings.Join(interleaved, ",")), nil

	case "CUSTOM":
		if spec.CustomFormula == "" {
			return "", errors.New("CUSTOM requires customFormula")
		}
		return replaceRowPlaceholder(spec.CustomFormula, row), nil

	default:
		return "", fmt.Errorf("unsupported column formula: %s (use VLOOKUP, IF, SUMIF, AVERAGEIF, CONCAT, CUSTOM)", spec.Func)
	}
}

// colNameToNumber converts "A" → 1, "AA" → 27.
func colNameToNumber(name string) int {
	n := 0
	for i := 0; i < len(name); i++ {
		n = n*26 + int(name[i]-'A') + 1
	}
	return n
}

// replaceRowPlaceholder substitutes {row} in a string with the row number.
func replaceRowPlaceholder(s string, row int) string {
	return strings.ReplaceAll(s, "{row}", strconv.Itoa(row))
}

func defaultStr(s, def string) string {
	if s == "" {
		return def
	}
	return s
}

// detectLastRow finds the last non-empty row across all columns in a sheet.
func detectLastRow(f *excelize.File, sheet string) int {
	rows, _ := f.GetRows(sheet)
	lastRow := 1
	for r, row := range rows {
		for _, cell := range row {
			if cell != "" {
				lastRow = r + 1
				break
			}
		}
	}
	if lastRow < 2 {
		lastRow = 2
	}
	return lastRow
}

// ---------------------------------------------------------------------------
// Feature 6: Excel ↔ JSON conversion
// ---------------------------------------------------------------------------

// JSONReadOptions controls how a sheet is exported to JSON.
type JSONReadOptions struct {
	Sheet       string `json:"sheet,omitempty"`
	Range       string `json:"range,omitempty"`
	WithHeaders bool   `json:"withHeaders,omitempty"`
	Typed       bool   `json:"typed,omitempty"`
	Limit       int    `json:"limit,omitempty"`
}

// ToJSON reads a sheet and returns it as a JSON-friendly structure.
// With WithHeaders=true and Typed=true, returns []map[string]any.
// With WithHeaders=true, returns []map[string]string.
// Otherwise returns [][]string.
func ToJSON(path string, opts JSONReadOptions) (any, error) {
	if opts.Typed {
		res, err := ReadTyped(path, ReadOptions{
			Sheet: opts.Sheet, Range: opts.Range, Limit: opts.Limit,
		})
		if err != nil {
			return nil, err
		}
		if opts.WithHeaders {
			return res.Rows, nil
		}
		out := make([][]any, 0, len(res.Rows))
		for _, r := range res.Rows {
			row := make([]any, 0, len(r))
			for _, v := range r {
				row = append(row, v)
			}
			out = append(out, row)
		}
		return out, nil
	}

	res, err := Read(path, ReadOptions{
		Sheet: opts.Sheet, Range: opts.Range, Limit: opts.Limit,
	})
	if err != nil {
		return nil, err
	}
	if opts.WithHeaders && len(res.Rows) > 0 {
		headers := res.Rows[0]
		records := make([]map[string]string, 0, len(res.Rows)-1)
		for _, row := range res.Rows[1:] {
			m := make(map[string]string, len(headers))
			for i, h := range headers {
				if i < len(row) {
					m[h] = row[i]
				}
			}
			records = append(records, m)
		}
		return records, nil
	}
	return res.Rows, nil
}

// FromJSON creates a workbook from JSON data.
// data can be []map[string]any, [][]any, or []any (JSON arrays).
func FromJSON(path string, sheetName string, data any, overwrite bool) (int, error) {
	if !overwrite {
		if _, err := os.Stat(path); err == nil {
			return 0, fmt.Errorf("file already exists (use --force to overwrite): %s", path)
		}
	}

	if sheetName == "" {
		sheetName = "Sheet1"
	}

	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	first := f.GetSheetName(0)
	if first != sheetName {
		if _, err := f.NewSheet(sheetName); err != nil {
			return 0, err
		}
		if err := f.DeleteSheet(first); err != nil {
			return 0, err
		}
	}

	rowCount := 0

	switch d := data.(type) {
	case []map[string]any:
		if len(d) == 0 {
			break
		}
		headerMap := map[string]int{}
		var headers []string
		for _, obj := range d {
			for k := range obj {
				if _, ok := headerMap[k]; !ok {
					headerMap[k] = len(headers)
					headers = append(headers, k)
				}
			}
		}
		for c, h := range headers {
			ref, _ := excelize.CoordinatesToCellName(c+1, 1)
			_ = f.SetCellValue(sheetName, ref, h)
		}
		rowCount = 1
		for _, obj := range d {
			rowCount++
			for c, h := range headers {
				ref, _ := excelize.CoordinatesToCellName(c+1, rowCount)
				if v, ok := obj[h]; ok {
					_ = f.SetCellValue(sheetName, ref, v)
				}
			}
		}

	case []any:
		if len(d) > 0 {
			if _, ok := d[0].(map[string]any); ok {
				objs := make([]map[string]any, len(d))
				for i, item := range d {
					objs[i], _ = item.(map[string]any)
				}
				return FromJSON(path, sheetName, objs, overwrite)
			}
		}
		for _, item := range d {
			if row, ok := item.([]any); ok {
				rowCount++
				for c, v := range row {
					ref, _ := excelize.CoordinatesToCellName(c+1, rowCount)
					_ = f.SetCellValue(sheetName, ref, v)
				}
			}
		}

	case [][]any:
		for _, row := range d {
			rowCount++
			for c, v := range row {
				ref, _ := excelize.CoordinatesToCellName(c+1, rowCount)
				_ = f.SetCellValue(sheetName, ref, v)
			}
		}

	default:
		return 0, fmt.Errorf("unsupported data type: %T (pass []map[string]any or [][]any)", data)
	}

	if err := f.SaveAs(path); err != nil {
		return 0, err
	}
	return rowCount, nil
}

// ---------------------------------------------------------------------------
// Feature 7: Range operations (fill, copy)
// ---------------------------------------------------------------------------

// FillRange fills every cell in the range with the given value.
func FillRange(path, sheet, cellRange string, value any) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	parts := strings.Split(cellRange, ":")
	if len(parts) != 2 {
		return fmt.Errorf("invalid range: %s (expected e.g. A1:C5)", cellRange)
	}
	startCol, startRow, _ := excelize.CellNameToCoordinates(parts[0])
	endCol, endRow, _ := excelize.CellNameToCoordinates(parts[1])

	for r := startRow; r <= endRow; r++ {
		for c := startCol; c <= endCol; c++ {
			ref, _ := excelize.CoordinatesToCellName(c, r)
			if err := f.SetCellValue(sheet, ref, value); err != nil {
				return err
			}
		}
	}
	return f.Save()
}

// CopyRange copies cell values from srcRange to dstRange (top-left anchored).
func CopyRange(path, sheet, srcRange, dstRange string) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	srcParts := strings.Split(srcRange, ":")
	if len(srcParts) != 2 {
		return fmt.Errorf("invalid src range: %s", srcRange)
	}
	srcStartCol, srcStartRow, _ := excelize.CellNameToCoordinates(srcParts[0])
	srcEndCol, srcEndRow, _ := excelize.CellNameToCoordinates(srcParts[1])

	dstCol, dstRow, _ := excelize.CellNameToCoordinates(dstRange)

	type cellData struct {
		value   any
		formula string
	}
	srcData := make([][]cellData, srcEndRow-srcStartRow+1)
	for r := srcStartRow; r <= srcEndRow; r++ {
		rowData := make([]cellData, srcEndCol-srcStartCol+1)
		for c := srcStartCol; c <= srcEndCol; c++ {
			ref, _ := excelize.CoordinatesToCellName(c, r)
			val, _ := f.GetCellValue(sheet, ref)
			formula, _ := f.GetCellFormula(sheet, ref)
			rowData[c-srcStartCol] = cellData{value: val, formula: formula}
		}
		srcData[r-srcStartRow] = rowData
	}

	for r, rowData := range srcData {
		for c, cd := range rowData {
			ref, _ := excelize.CoordinatesToCellName(dstCol+c, dstRow+r)
			if cd.formula != "" {
				_ = f.SetCellFormula(sheet, ref, cd.formula)
			} else {
				_ = f.SetCellValue(sheet, ref, cd.value)
			}
		}
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Feature 8: Multi-series chart
// ---------------------------------------------------------------------------

// MultiChartSpec describes a chart with multiple data series.
type MultiChartSpec struct {
	Type   string       `json:"type"` // col, bar, line, area, pie, scatter
	Cell   string       `json:"cell"` // anchor cell
	Title  string       `json:"title,omitempty"`
	Width  int          `json:"width,omitempty"`  // default 12
	Height int          `json:"height,omitempty"` // default 8
	Series []SeriesSpec `json:"series"`           // required: at least 1
}

// SeriesSpec describes one data series in a chart.
type SeriesSpec struct {
	Name     string `json:"name,omitempty"`     // legend label
	CatRange string `json:"catRange,omitempty"` // category (X) range
	ValRange string `json:"valRange"`           // value (Y) range
}

// AddMultiChart creates a chart with multiple data series.
func AddMultiChart(path, sheet string, spec MultiChartSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	width := uint(spec.Width)
	if width <= 0 {
		width = 12
	}
	height := uint(spec.Height)
	if height <= 0 {
		height = 8
	}

	var chartType excelize.ChartType
	switch strings.ToLower(spec.Type) {
	case "col", "column":
		chartType = excelize.Col
	case "bar":
		chartType = excelize.Bar
	case "line":
		chartType = excelize.Line
	case "area":
		chartType = excelize.Area
	case "pie":
		chartType = excelize.Pie
	case "scatter":
		chartType = excelize.Scatter
	default:
		return fmt.Errorf("unsupported chart type: %s", spec.Type)
	}

	chart := &excelize.Chart{
		Type: chartType,
		Dimension: excelize.ChartDimension{
			Width:  width,
			Height: height,
		},
	}
	if spec.Title != "" {
		chart.Title = []excelize.RichTextRun{{Text: spec.Title}}
	}

	for _, s := range spec.Series {
		cs := excelize.ChartSeries{
			Values: fmt.Sprintf("'%s'!%s", sheet, s.ValRange),
		}
		if s.Name != "" {
			cs.Name = s.Name
		}
		if s.CatRange != "" {
			cs.Categories = fmt.Sprintf("'%s'!%s", sheet, s.CatRange)
		}
		chart.Series = append(chart.Series, cs)
	}

	return f.AddChart(sheet, spec.Cell, chart)
}

// ---------------------------------------------------------------------------
// Feature 9: Enhanced conditional formatting (data bars, icon sets)
// ---------------------------------------------------------------------------

// DataBarSpec describes a data bar conditional format.
type DataBarSpec struct {
	Range     string `json:"range"`
	MinColor  string `json:"minColor,omitempty"`  // hex, default "638EC6"
	MaxColor  string `json:"maxColor,omitempty"`  // hex, default "638EC6"
	ShowValue *bool  `json:"showValue,omitempty"` // default: true
}

// AddDataBar applies a data bar conditional format to a range.
func AddDataBar(path, sheet string, spec DataBarSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	minColor := spec.MinColor
	if minColor == "" {
		minColor = "638EC6"
	}
	maxColor := spec.MaxColor
	if maxColor == "" {
		maxColor = "638EC6"
	}
	barOnly := false
	if spec.ShowValue != nil {
		barOnly = !*spec.ShowValue
	}

	format := []excelize.ConditionalFormatOptions{
		{
			Type:     "data_bar",
			Criteria: "=",
			MinType:  "min",
			MaxType:  "max",
			MinColor: minColor,
			MaxColor: maxColor,
			BarColor: maxColor,
			BarOnly:  barOnly,
		},
	}
	if err := f.SetConditionalFormat(sheet, spec.Range, format); err != nil {
		return err
	}
	return f.Save()
}

// IconSetSpec describes an icon set conditional format.
type IconSetSpec struct {
	Range   string `json:"range"`
	Style   string `json:"style,omitempty"` // 3Arrows, 3TrafficLights, 3Stars, etc.
	Reverse bool   `json:"reverse,omitempty"`
}

// AddIconSet applies an icon set conditional format to a range.
func AddIconSet(path, sheet string, spec IconSetSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	style := "3Arrows"
	if spec.Style != "" {
		style = spec.Style
	}

	format := []excelize.ConditionalFormatOptions{
		{
			Type:         "icon_set",
			IconStyle:    style,
			ReverseIcons: spec.Reverse,
		},
	}
	if err := f.SetConditionalFormat(sheet, spec.Range, format); err != nil {
		return err
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Feature 10: Color scale conditional formatting
// ---------------------------------------------------------------------------

// ColorScaleSpec describes a color scale conditional format.
type ColorScaleSpec struct {
	Range    string `json:"range"`
	MinColor string `json:"minColor,omitempty"` // hex, default "F8696B"
	MaxColor string `json:"maxColor,omitempty"` // hex, default "63BE7B"
	MidColor string `json:"midColor,omitempty"` // hex (omit for 2-color scale)
	MinType  string `json:"minType,omitempty"`  // min (default), num, percent, percentile, formula
	MaxType  string `json:"maxType,omitempty"`  // max (default), num, percent, percentile, formula
	MidType  string `json:"midType,omitempty"`  // num (default), percent, percentile, formula
	MinValue string `json:"minValue,omitempty"` // default "0"
	MaxValue string `json:"maxValue,omitempty"` // default "0"
	MidValue string `json:"midValue,omitempty"` // default "50"
}

// AddColorScale applies a 2-color or 3-color scale conditional format to a range.
func AddColorScale(path, sheet string, spec ColorScaleSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}

	minColor := defaultStr(spec.MinColor, "F8696B")
	maxColor := defaultStr(spec.MaxColor, "63BE7B")
	minType := defaultStr(spec.MinType, "min")
	maxType := defaultStr(spec.MaxType, "max")
	midType := defaultStr(spec.MidType, "num")
	minValue := defaultStr(spec.MinValue, "0")
	maxValue := defaultStr(spec.MaxValue, "0")
	midValue := defaultStr(spec.MidValue, "50")

	cfType := "2_color_scale"
	format := []excelize.ConditionalFormatOptions{
		{
			Type:     cfType,
			Criteria: "=",
			MinType:  minType,
			MaxType:  maxType,
			MinColor: minColor,
			MaxColor: maxColor,
			MinValue: minValue,
			MaxValue: maxValue,
		},
	}
	if spec.MidColor != "" {
		format[0].Type = "3_color_scale"
		format[0].MidType = midType
		format[0].MidColor = spec.MidColor
		format[0].MidValue = midValue
	}
	if err := f.SetConditionalFormat(sheet, spec.Range, format); err != nil {
		return err
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Feature 11: Hyperlinks
// ---------------------------------------------------------------------------

// HyperlinkSpec describes a hyperlink to set on a cell.
type HyperlinkSpec struct {
	Cell    string `json:"cell"`              // e.g. "A1"
	Link    string `json:"link"`              // URL or "Sheet1!A1" for internal
	Type    string `json:"type,omitempty"`    // "External" (default) or "Location"
	Display string `json:"display,omitempty"` // display text (default: current cell value)
	Tooltip string `json:"tooltip,omitempty"` // hover tooltip
}

// HyperlinkResult is the output of GetHyperlink.
type HyperlinkResult struct {
	Cell   string `json:"cell"`
	Link   string `json:"link"`
	Exists bool   `json:"exists"`
}

// SetHyperlink sets a hyperlink on a cell. Applies default link style (blue + underline).
func SetHyperlink(path, sheet string, spec HyperlinkSpec) error {
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	if !sheetExists(f, sheet) {
		return fmt.Errorf("sheet not found: %s", sheet)
	}
	linkType := spec.Type
	if linkType == "" {
		linkType = "External"
	}
	var opts []excelize.HyperlinkOpts
	if spec.Display != "" || spec.Tooltip != "" {
		opt := excelize.HyperlinkOpts{}
		if spec.Display != "" {
			opt.Display = &spec.Display
		}
		if spec.Tooltip != "" {
			opt.Tooltip = &spec.Tooltip
		}
		opts = append(opts, opt)
	}
	if err := f.SetCellHyperLink(sheet, spec.Cell, spec.Link, linkType, opts...); err != nil {
		return fmt.Errorf("set hyperlink: %w", err)
	}
	// If display text is set, write it to the cell.
	if spec.Display != "" {
		if err := f.SetCellValue(sheet, spec.Cell, spec.Display); err != nil {
			return err
		}
	}
	// Apply default hyperlink style (blue underline).
	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "1265BE", Underline: "single"},
	})
	if err == nil {
		_ = f.SetCellStyle(sheet, spec.Cell, spec.Cell, styleID)
	}
	return f.Save()
}

// GetHyperlink reads the hyperlink for a cell.
func GetHyperlink(path, sheet, cell string) (*HyperlinkResult, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	exists, link, err := f.GetCellHyperLink(sheet, cell)
	if err != nil {
		return nil, err
	}
	return &HyperlinkResult{Cell: cell, Link: link, Exists: exists}, nil
}

// ---------------------------------------------------------------------------
// Default Font
// ---------------------------------------------------------------------------

// GetDefaultFont reads the workbook's default font name.
// Excel's built-in default is "Calibri".
func GetDefaultFont(path string) (string, error) {
	f, err := open(path)
	if err != nil {
		return "", err
	}
	defer func() { _ = f.Close() }()
	name, err := f.GetDefaultFont()
	if err != nil {
		return "", fmt.Errorf("get default font: %w", err)
	}
	return name, nil
}

// SetDefaultFont changes the workbook's default font name.
// This affects cells that have no explicit font style.
func SetDefaultFont(path, fontName string) error {
	if fontName == "" {
		return fmt.Errorf("font name is required")
	}
	f, err := open(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()
	if err := f.SetDefaultFont(fontName); err != nil {
		return fmt.Errorf("set default font: %w", err)
	}
	return f.Save()
}

// ---------------------------------------------------------------------------
// Cell Style Read
// ---------------------------------------------------------------------------

// CellStyleInfo represents the style properties of a cell.
type CellStyleInfo struct {
	Cell         string  `json:"cell"`
	Bold         bool    `json:"bold,omitempty"`
	Italic       bool    `json:"italic,omitempty"`
	Underline    bool    `json:"underline,omitempty"`
	Strike       bool    `json:"strike,omitempty"`
	FontSize     float64 `json:"fontSize,omitempty"`
	FontFamily   string  `json:"fontFamily,omitempty"`
	FontColor    string  `json:"fontColor,omitempty"`
	BgColor      string  `json:"bgColor,omitempty"`
	NumberFormat string  `json:"numberFormat,omitempty"`
	Align        string  `json:"align,omitempty"`
	Valign       string  `json:"valign,omitempty"`
	WrapText     bool    `json:"wrapText,omitempty"`
	TextRotation int     `json:"textRotation,omitempty"`
	Indent       int     `json:"indent,omitempty"`
	ShrinkToFit  bool    `json:"shrinkToFit,omitempty"`
	StyleIndex   int     `json:"styleIndex"`
}

// GetCellStyleInfo reads the style properties of a cell.
// Returns nil style info if the cell has the default style (index 0).
func GetCellStyleInfo(path, sheet, cell string) (*CellStyleInfo, error) {
	f, err := open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()
	if sheet == "" {
		sheet = f.GetSheetName(0)
	}
	idx, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return nil, fmt.Errorf("get cell style: %w", err)
	}
	// Retrieve the full style definition.
	style, err := f.GetStyle(idx)
	if err != nil {
		return nil, fmt.Errorf("get style definition: %w", err)
	}
	info := &CellStyleInfo{
		Cell:       cell,
		StyleIndex: idx,
	}
	if style.Font != nil {
		info.Bold = style.Font.Bold
		info.Italic = style.Font.Italic
		info.Underline = style.Font.Underline != ""
		info.Strike = style.Font.Strike
		if style.Font.Size > 0 {
			info.FontSize = style.Font.Size
		}
		info.FontFamily = style.Font.Family
		info.FontColor = style.Font.Color
	}
	// Fill: Pattern 1 = solid fill, Color[0] is the background.
	if style.Fill.Pattern == 1 && len(style.Fill.Color) > 0 {
		info.BgColor = style.Fill.Color[0]
	}
	if style.CustomNumFmt != nil && *style.CustomNumFmt != "" {
		info.NumberFormat = *style.CustomNumFmt
	} else if style.NumFmt > 0 {
		info.NumberFormat = fmt.Sprintf("%d", style.NumFmt)
	}
	if style.Alignment != nil {
		info.Align = style.Alignment.Horizontal
		info.Valign = style.Alignment.Vertical
		info.WrapText = style.Alignment.WrapText
		info.TextRotation = style.Alignment.TextRotation
		info.Indent = style.Alignment.Indent
		info.ShrinkToFit = style.Alignment.ShrinkToFit
	}
	return info, nil
}
