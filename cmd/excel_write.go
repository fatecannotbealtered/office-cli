package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/fatecannotbealtered/office-cli/internal/engine/excel"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// Excel write commands: write, append, create, to-csv, from-csv, from-json,
// formula, column-formula, copy, fill-range, copy-range
// ---------------------------------------------------------------------------

var excelWriteCmd = &cobra.Command{
	Use:   "write <FILE>",
	Short: "Write one or many cells (single via --ref/--value, batch via --cells JSON)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		ref, _ := cmd.Flags().GetString("ref")
		value, _ := cmd.Flags().GetString("value")
		formula, _ := cmd.Flags().GetString("formula")
		cellsRaw, _ := cmd.Flags().GetString("cells")

		var cells []excel.CellWrite
		if cellsRaw != "" {
			data, err := readSpecArg(cellsRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &cells); err != nil {
				return emitError("invalid --cells JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}
		if ref != "" {
			cells = append(cells, excel.CellWrite{Sheet: sheet, Ref: ref, Value: value, Formula: formula})
		}
		if len(cells) == 0 {
			return emitError("provide --ref/--value or --cells", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("write cells", map[string]any{"file": path, "count": len(cells)}) {
			return nil
		}

		if err := excel.WriteCells(path, cells); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "written": len(cells)})
			return nil
		}
		output.Success(fmt.Sprintf("wrote %d cell(s) to %s", len(cells), path))
		return nil
	},
}

var excelAppendCmd = &cobra.Command{
	Use:   "append <FILE>",
	Short: "Append rows to a sheet",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		rowsRaw, _ := cmd.Flags().GetString("rows")
		if rowsRaw == "" {
			return emitError("--rows is required (JSON array of arrays, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(rowsRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var rows [][]any
		if err := json.Unmarshal(data, &rows); err != nil {
			return emitError("invalid --rows JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("append rows", map[string]any{"file": path, "sheet": sheet, "count": len(rows)}) {
			return nil
		}

		startRow, err := excel.AppendRows(path, sheet, rows)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheet": sheet, "appended": len(rows), "startRow": startRow})
			return nil
		}
		output.Success(fmt.Sprintf("appended %d row(s) starting at row %d in %s", len(rows), startRow, path))
		return nil
	},
}

var excelCreateCmd = &cobra.Command{
	Use:   "create <FILE>",
	Short: "Create a new .xlsx workbook from a JSON spec",
	Long: `Create a new .xlsx workbook from a JSON spec. Spec shape:
  {
    "sheets": [
      { "name": "Sales", "rows": [["A","B"], [1,2]] },
      { "name": "Notes", "rows": [["hello"]] }
    ]
  }`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		specRaw, _ := cmd.Flags().GetString("spec")
		if specRaw == "" {
			return emitError("--spec is required (JSON object, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var spec excel.CreateSpec
		if err := json.Unmarshal(data, &spec); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("create workbook", map[string]any{"file": path, "sheets": len(spec.Sheets)}) {
			return nil
		}

		if err := excel.Create(path, spec, forceMode); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheets": len(spec.Sheets)})
			return nil
		}
		output.Success(fmt.Sprintf("created %s with %d sheet(s)", path, len(spec.Sheets)))
		return nil
	},
}

var excelToCSVCmd = &cobra.Command{
	Use:   "to-csv <FILE>",
	Short: "Export one sheet as CSV",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		dest, _ := cmd.Flags().GetString("output")
		bom, _ := cmd.Flags().GetBool("bom")
		if dest == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		dest, err = resolveOutput(dest)
		if err != nil {
			return err
		}

		if dryRunOutput("export to CSV", map[string]any{"file": path, "sheet": sheet, "output": dest, "bom": bom}) {
			return nil
		}

		if err := excel.ToCSV(path, sheet, dest, bom); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": dest, "bom": bom})
			return nil
		}
		output.Success(fmt.Sprintf("exported %s -> %s", path, dest))
		return nil
	},
}

var excelFromCSVCmd = &cobra.Command{
	Use:   "from-csv <CSV>",
	Short: "Build a new .xlsx workbook from a CSV file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		csvPath, err := resolveInput(args[0], "csv", "txt")
		if err != nil {
			return err
		}
		out, _ := cmd.Flags().GetString("output")
		sheet, _ := cmd.Flags().GetString("sheet")
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("import CSV to xlsx", map[string]any{"csv": csvPath, "output": out, "sheet": sheet}) {
			return nil
		}

		rows, err := excel.FromCSV(csvPath, out, sheet, forceMode)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, csvPath, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "csv": csvPath, "output": out, "sheet": sheet, "rows": rows})
			return nil
		}
		output.Success(fmt.Sprintf("imported %d row(s) %s -> %s", rows, csvPath, out))
		return nil
	},
}

var excelFormulaCmd = &cobra.Command{
	Use:   "formula <FILE>",
	Short: "Generate formulas for columns (SUM, AVERAGE, COUNT, COUNTIF, SUMIF, etc.)",
	Long: `Generate formulas for columns. Single mode uses --col and --func;
batch mode uses --spec with a JSON array. With --apply, formulas are written
directly to the workbook.`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		specRaw, _ := cmd.Flags().GetString("spec")
		apply, _ := cmd.Flags().GetBool("apply")

		var results []excel.FormulaResult

		if specRaw != "" {
			// Batch mode
			data, err := readSpecArg(specRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			var specs []excel.FormulaSpec
			if err := json.Unmarshal(data, &specs); err != nil {
				return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			results, err = excel.GenerateFormulas(path, sheet, specs)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
		} else {
			// Single mode
			col, _ := cmd.Flags().GetString("col")
			fn, _ := cmd.Flags().GetString("func")
			if col == "" || fn == "" {
				return emitError("--col and --func are required (or use --spec for batch mode)", output.ErrValidation, "", ExitBadArgs)
			}
			dataRange, _ := cmd.Flags().GetString("data-range")
			outputCell, _ := cmd.Flags().GetString("output-cell")
			criteria, _ := cmd.Flags().GetString("criteria")

			spec := excel.FormulaSpec{
				Column:     col,
				Func:       fn,
				DataRange:  dataRange,
				OutputCell: outputCell,
				Criteria:   criteria,
			}
			r, err := excel.GenerateFormula(path, sheet, spec)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			results = []excel.FormulaResult{*r}
		}

		if apply {
			if dryRunOutput("apply formulas", map[string]any{"file": path, "count": len(results)}) {
				return nil
			}
			if err := excel.ApplyFormulas(path, sheet, results); err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
		}

		if jsonMode {
			payload := map[string]any{"file": path, "formulas": results}
			if apply {
				payload["applied"] = true
			}
			output.PrintJSON(payload)
			return nil
		}

		for _, r := range results {
			applied := ""
			if apply {
				applied = " [applied]"
			}
			output.Info(fmt.Sprintf("%s = %s  (range: %s, %d rows)%s", r.Cell, r.Formula, r.DataRange, r.DataRows, applied))
		}
		if apply {
			output.Success(fmt.Sprintf("applied %d formula(s) to %s", len(results), path))
		}
		return nil
	},
}

var excelCopyCmd = &cobra.Command{
	Use:   "copy <FILE>",
	Short: "Deep-copy a workbook (all sheets, data, styles, charts, images)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		src, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		dst, _ := cmd.Flags().GetString("output")
		force, _ := cmd.Flags().GetBool("force")
		if dst == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		dst, err = resolveOutput(dst)
		if err != nil {
			return err
		}

		if dryRunOutput("copy workbook", map[string]any{"src": src, "dst": dst}) {
			return nil
		}

		if err := excel.CopyWorkbook(src, dst, force); err != nil {
			return emitError(err.Error(), output.ErrEngine, src, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "src": src, "dst": dst})
			return nil
		}
		output.Success(fmt.Sprintf("copied %s -> %s", src, dst))
		return nil
	},
}

var excelColumnFormulaCmd = &cobra.Command{
	Use:   "column-formula <FILE>",
	Short: "Apply row-level formulas (VLOOKUP, IF, SUMIF, AVERAGEIF, CONCAT, CUSTOM) down a column",
	Long: `Generates and optionally applies formulas per row. Supports VLOOKUP, IF,
SUMIF, AVERAGEIF, CONCAT, and custom formulas with {row} placeholder.

Use --spec for a JSON object, or individual flags. With --apply, formulas are
written to every row in the range.`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		specRaw, _ := cmd.Flags().GetString("spec")
		apply, _ := cmd.Flags().GetBool("apply")

		var spec excel.ColumnFormulaSpec

		if specRaw != "" {
			data, err := readSpecArg(specRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &spec); err != nil {
				return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		} else {
			col, _ := cmd.Flags().GetString("col")
			fn, _ := cmd.Flags().GetString("func")
			if col == "" || fn == "" {
				return emitError("--col and --func are required (or use --spec)", output.ErrValidation, "", ExitBadArgs)
			}
			spec = excel.ColumnFormulaSpec{
				Column: col,
				Func:   fn,
			}
			spec.StartRow, _ = cmd.Flags().GetInt("start-row")
			spec.EndRow, _ = cmd.Flags().GetInt("end-row")
		}

		if apply {
			if dryRunOutput("apply column formulas", map[string]any{"file": path, "col": spec.Column, "func": spec.Func}) {
				return nil
			}
			count, err := excel.ApplyColumnFormulas(path, sheet, spec)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"status": "ok", "file": path, "applied": count, "func": spec.Func})
				return nil
			}
			output.Success(fmt.Sprintf("applied %d %s formulas to %s", count, spec.Func, path))
			return nil
		}

		// Preview mode: generate without applying.
		result, err := excel.GenerateColumnFormula(path, sheet, spec)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "column": result.Column, "formula": result.Formula, "rowCount": result.RowCount, "sampleRow": result.SampleRow})
			return nil
		}
		output.Info(fmt.Sprintf("Column %s, rows %d-%d: %s", result.Column, result.SampleRow, result.SampleRow+result.RowCount-1, result.Formula))
		return nil
	},
}

var excelFromJSONCmd = &cobra.Command{
	Use:   "from-json <FILE>",
	Short: "Create a new .xlsx workbook from JSON data",
	Long: `Create a workbook from JSON data. Data can be an array of objects
([{header:value,...}]) or an array of arrays ([[v1,v2,...]]).`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		sheet, _ := cmd.Flags().GetString("sheet")
		dataRaw, _ := cmd.Flags().GetString("data")
		if dataRaw == "" {
			return emitError("--data is required (JSON array, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(dataRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		var parsed any
		if err := json.Unmarshal(data, &parsed); err != nil {
			return emitError("invalid --data JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("create from JSON", map[string]any{"file": path, "sheet": sheet}) {
			return nil
		}

		rows, err := excel.FromJSON(path, sheet, parsed, forceMode)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheet": sheet, "rows": rows})
			return nil
		}
		output.Success(fmt.Sprintf("created %s with %d row(s)", path, rows))
		return nil
	},
}

var excelFillRangeCmd = &cobra.Command{
	Use:   "fill-range <FILE>",
	Short: "Fill every cell in a range with a value",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		value, _ := cmd.Flags().GetString("value")
		if cellRange == "" {
			return emitError("--range is required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("fill range", map[string]any{"file": path, "range": cellRange, "value": value}) {
			return nil
		}

		if err := excel.FillRange(path, sheet, cellRange, value); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("filled %s in %s", cellRange, path))
		return nil
	},
}

var excelCopyRangeCmd = &cobra.Command{
	Use:   "copy-range <FILE>",
	Short: "Copy cell values from a source range to a destination (top-left anchored)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		src, _ := cmd.Flags().GetString("src")
		dst, _ := cmd.Flags().GetString("dst")
		if src == "" || dst == "" {
			return emitError("--src and --dst are required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("copy range", map[string]any{"file": path, "src": src, "dst": dst}) {
			return nil
		}

		if err := excel.CopyRange(path, sheet, src, dst); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "src": src, "dst": dst})
			return nil
		}
		output.Success(fmt.Sprintf("copied %s → %s in %s", src, dst, path))
		return nil
	},
}

func init() {
	excelWriteCmd.Flags().String("sheet", "", "Sheet name (default: first sheet, created if missing)")
	excelWriteCmd.Flags().String("ref", "", "Single-cell write target, e.g. B2 (use with --value)")
	excelWriteCmd.Flags().String("value", "", "Value to write at --ref")
	excelWriteCmd.Flags().String("formula", "", "Formula to write at --ref (e.g. =SUM(A1:A10))")
	excelWriteCmd.Flags().String("cells", "", "JSON array of {sheet,ref,value,formula} for batch writes (or @file.json)")
	excelCmd.AddCommand(excelWriteCmd)

	excelAppendCmd.Flags().String("sheet", "", "Sheet name (default: first sheet, created if missing)")
	excelAppendCmd.Flags().String("rows", "", "JSON array of arrays (or @file.json)")
	excelCmd.AddCommand(excelAppendCmd)

	excelCreateCmd.Flags().String("spec", "", "JSON spec describing sheets and rows (or @file.json)")
	excelCmd.AddCommand(excelCreateCmd)

	excelToCSVCmd.Flags().String("sheet", "", "Sheet to export (default: first sheet)")
	excelToCSVCmd.Flags().String("output", "", "Output CSV path (required)")
	excelToCSVCmd.Flags().Bool("bom", true, "Prepend a UTF-8 BOM (recommended for Excel on Windows)")
	excelCmd.AddCommand(excelToCSVCmd)

	excelFromCSVCmd.Flags().String("output", "", "Output xlsx path (required)")
	excelFromCSVCmd.Flags().String("sheet", "Sheet1", "Sheet name in the new workbook")
	excelCmd.AddCommand(excelFromCSVCmd)

	excelFormulaCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelFormulaCmd.Flags().String("col", "", "Column letter or number (e.g. B or 2) — required in single mode")
	excelFormulaCmd.Flags().String("func", "", "Function: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, COUNTIF, SUMIF — required in single mode")
	excelFormulaCmd.Flags().String("data-range", "", "Data range (e.g. B2:B100) — auto-detected if empty")
	excelFormulaCmd.Flags().String("output-cell", "", "Output cell (e.g. B101) — auto-detected if empty")
	excelFormulaCmd.Flags().String("criteria", "", "Criteria for COUNTIF/SUMIF (e.g. \">100\")")
	excelFormulaCmd.Flags().String("spec", "", "Batch formula spec: JSON array of {column,func,...} objects (or @file.json)")
	excelFormulaCmd.Flags().Bool("apply", false, "Write formula(s) directly to the workbook")
	excelCmd.AddCommand(excelFormulaCmd)

	excelCopyCmd.Flags().String("output", "", "Output file path (required)")
	excelCopyCmd.Flags().Bool("force", false, "Overwrite output file if it exists")
	excelCmd.AddCommand(excelCopyCmd)

	excelColumnFormulaCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelColumnFormulaCmd.Flags().String("col", "", "Target column (e.g. D) — required")
	excelColumnFormulaCmd.Flags().String("func", "", "Function: VLOOKUP, IF, SUMIF, AVERAGEIF, CONCAT, CUSTOM — required")
	excelColumnFormulaCmd.Flags().Int("start-row", 2, "First data row (default: 2)")
	excelColumnFormulaCmd.Flags().Int("end-row", 0, "Last data row (auto-detected if 0)")
	excelColumnFormulaCmd.Flags().String("spec", "", "JSON spec for column formula (or @file.json)")
	excelColumnFormulaCmd.Flags().Bool("apply", false, "Write formulas to all rows")
	excelCmd.AddCommand(excelColumnFormulaCmd)

	excelToJSONCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelToJSONCmd.Flags().String("range", "", "Cell range (e.g. A1:D10)")
	excelToJSONCmd.Flags().Bool("with-headers", false, "Treat first row as keys")
	excelToJSONCmd.Flags().Bool("typed", false, "Return native types")
	excelToJSONCmd.Flags().Int("limit", 0, "Max rows (0 = all)")
	excelCmd.AddCommand(excelToJSONCmd)

	excelFromJSONCmd.Flags().String("sheet", "Sheet1", "Sheet name in the new workbook")
	excelFromJSONCmd.Flags().String("data", "", "JSON data: array of objects or arrays (or @file.json)")
	excelCmd.AddCommand(excelFromJSONCmd)

	excelFillRangeCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelFillRangeCmd.Flags().String("range", "", "Cell range (e.g. A1:C5) — required")
	excelFillRangeCmd.Flags().String("value", "", "Value to fill — required")
	excelCmd.AddCommand(excelFillRangeCmd)

	excelCopyRangeCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelCopyRangeCmd.Flags().String("src", "", "Source range (e.g. A1:C5) — required")
	excelCopyRangeCmd.Flags().String("dst", "", "Destination top-left cell (e.g. E1) — required")
	excelCmd.AddCommand(excelCopyRangeCmd)

	markWrite(excelWriteCmd)
	markWrite(excelAppendCmd)
	markWrite(excelCreateCmd)
	markWrite(excelToCSVCmd)
	markWrite(excelFromCSVCmd)
	markWrite(excelFormulaCmd)
	markWrite(excelCopyCmd)
	markWrite(excelColumnFormulaCmd)
	markWrite(excelFromJSONCmd)
	markWrite(excelFillRangeCmd)
	markWrite(excelCopyRangeCmd)
}
