package cmd

import (
	"fmt"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/excel"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var excelCmd = &cobra.Command{
	Use:   "excel",
	Short: "Read, write, search and convert .xlsx workbooks",
}

func init() {
	rootCmd.AddCommand(excelCmd)

	excelCmd.AddCommand(excelSheetsCmd)

	excelReadCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelReadCmd.Flags().String("range", "", "Cell range, e.g. A1:D10 (default: whole sheet)")
	excelReadCmd.Flags().Int("limit", 0, "Max rows to return (0 = all)")
	excelReadCmd.Flags().Bool("with-headers", false, "Treat first row as headers; output objects instead of arrays")
	excelReadCmd.Flags().Bool("typed", false, "Return typed values (numbers as numbers, bools as bools) with type annotations")
	excelCmd.AddCommand(excelReadCmd)

	excelSearchCmd.Flags().String("sheet", "", "Restrict search to this sheet (default: all sheets)")
	excelSearchCmd.Flags().Int("limit", 100, "Max matches to return")
	excelCmd.AddCommand(excelSearchCmd)

	excelCellCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelCmd.AddCommand(excelCellCmd)

	excelInfoCmd.Flags().Int("preview", 0, "Include the first N rows of every sheet")
	excelCmd.AddCommand(excelInfoCmd)
}

// ---------------------------------------------------------------------------
// Read commands
// ---------------------------------------------------------------------------

var excelSheetsCmd = &cobra.Command{
	Use:   "sheets <FILE>",
	Short: "List worksheets with row/col counts",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheets, err := excel.ListSheets(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			flat := make([]output.FlatSheet, 0, len(sheets))
			for _, s := range sheets {
				flat = append(flat, excel.MapSheetInfo(s))
			}
			output.PrintJSON(flat)
			return nil
		}

		rows := make([][]string, 0, len(sheets))
		for _, s := range sheets {
			hidden := ""
			if s.Hidden {
				hidden = "yes"
			}
			rows = append(rows, []string{
				fmt.Sprintf("%d", s.Index),
				s.Name,
				fmt.Sprintf("%d", s.Rows),
				fmt.Sprintf("%d", s.Cols),
				s.Dimension,
				hidden,
			})
		}
		output.Table([]string{"index", "name", "rows", "cols", "dimension", "hidden"}, rows)
		return nil
	},
}

var excelReadCmd = &cobra.Command{
	Use:   "read <FILE>",
	Short: "Read sheet contents (full sheet or A1:D10 range)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		rng, _ := cmd.Flags().GetString("range")
		limit, _ := cmd.Flags().GetInt("limit")
		withHeaders, _ := cmd.Flags().GetBool("with-headers")
		typed, _ := cmd.Flags().GetBool("typed")

		// Typed mode: return values with native Go types and type annotations.
		if typed {
			res, err := excel.ReadTyped(path, excel.ReadOptions{Sheet: sheet, Range: rng, Limit: limit})
			if err != nil {
				if strings.Contains(err.Error(), "sheet not found") {
					return emitError(err.Error(), output.ErrNotFound, path, ExitNotFound)
				}
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(res)
				return nil
			}
			// Human mode: print as table with type row.
			if len(res.Rows) == 0 {
				output.Gray("(empty sheet)")
				return nil
			}
			hdrs := []string{"#"}
			hdrs = append(hdrs, res.Headers...)
			typeRow := []string{"type"}
			for _, h := range res.Headers {
				typeRow = append(typeRow, res.Types[h])
			}
			tblRows := [][]string{typeRow}
			for i, r := range res.Rows {
				row := []string{fmt.Sprintf("%d", i+1)}
				for _, h := range res.Headers {
					row = append(row, fmt.Sprintf("%v", r[h]))
				}
				tblRows = append(tblRows, row)
			}
			output.Table(hdrs, tblRows)
			return nil
		}

		res, err := excel.Read(path, excel.ReadOptions{Sheet: sheet, Range: rng, Limit: limit})
		if err != nil {
			if strings.Contains(err.Error(), "sheet not found") {
				return emitError(err.Error(), output.ErrNotFound, path, ExitNotFound)
			}
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			if withHeaders && len(res.Rows) > 0 {
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
				output.PrintJSON(map[string]any{
					"sheet":   res.Sheet,
					"range":   res.Range,
					"records": records,
				})
				return nil
			}
			output.PrintJSON(res)
			return nil
		}

		if len(res.Rows) == 0 {
			output.Gray("(empty sheet)")
			return nil
		}
		headers := []string{"#"}
		for i := 1; i <= len(res.Rows[0]); i++ {
			headers = append(headers, columnLetter(i))
		}
		rows := make([][]string, 0, len(res.Rows))
		for i, r := range res.Rows {
			rows = append(rows, append([]string{fmt.Sprintf("%d", i+1)}, r...))
		}
		output.Table(headers, rows)
		return nil
	},
}

var excelSearchCmd = &cobra.Command{
	Use:   "search <FILE> <KEYWORD>",
	Short: "Find cells containing keyword (case-insensitive)",
	Args:  cobra.ExactArgs(2),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		keyword := args[1]
		sheet, _ := cmd.Flags().GetString("sheet")
		limit, _ := cmd.Flags().GetInt("limit")

		matches, err := excel.Search(path, keyword, sheet, limit)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{
				"file":    path,
				"keyword": keyword,
				"count":   len(matches),
				"matches": matches,
			})
			return nil
		}

		if len(matches) == 0 {
			output.Gray(fmt.Sprintf("no matches for %q", keyword))
			return nil
		}
		rows := make([][]string, 0, len(matches))
		for _, m := range matches {
			rows = append(rows, []string{m.Sheet, m.Ref, m.Value, m.Formula})
		}
		output.Table([]string{"sheet", "ref", "value", "formula"}, rows)
		return nil
	},
}

var excelCellCmd = &cobra.Command{
	Use:   "cell <FILE> <REF>",
	Short: "Read a single cell value (e.g. B2). Lightweight; avoids loading the full sheet.",
	Args:  cobra.ExactArgs(2),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		ref := args[1]
		sheet, _ := cmd.Flags().GetString("sheet")

		val, formula, err := excel.ReadCell(path, sheet, ref)
		if err != nil {
			if strings.Contains(err.Error(), "sheet not found") {
				return emitError(err.Error(), output.ErrNotFound, path, ExitNotFound)
			}
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			payload := map[string]any{"file": path, "sheet": sheet, "ref": ref, "value": val}
			if formula != "" {
				payload["formula"] = formula
			}
			output.PrintJSON(payload)
			return nil
		}
		fmt.Println(val)
		return nil
	},
}

var excelInfoCmd = &cobra.Command{
	Use:   "info <FILE>",
	Short: "Show a quick overview: sheet names, dimensions, optional preview rows",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		previewRows, _ := cmd.Flags().GetInt("preview")

		sheets, err := excel.ListSheets(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		flat := make([]map[string]any, 0, len(sheets))
		for _, s := range sheets {
			entry := map[string]any{
				"name":      s.Name,
				"index":     s.Index,
				"rows":      s.Rows,
				"cols":      s.Cols,
				"dimension": s.Dimension,
				"hidden":    s.Hidden,
			}
			if previewRows > 0 {
				res, err := excel.Read(path, excel.ReadOptions{Sheet: s.Name, Limit: previewRows})
				if err == nil {
					entry["preview"] = res.Rows
				}
			}
			flat = append(flat, entry)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "sheets": flat})
			return nil
		}
		rows := make([][]string, 0, len(sheets))
		for _, s := range sheets {
			rows = append(rows, []string{
				fmt.Sprintf("%d", s.Index),
				s.Name,
				fmt.Sprintf("%d", s.Rows),
				fmt.Sprintf("%d", s.Cols),
				s.Dimension,
			})
		}
		output.Table([]string{"#", "name", "rows", "cols", "dimension"}, rows)
		return nil
	},
}

var excelToJSONCmd = &cobra.Command{
	Use:   "to-json <FILE>",
	Short: "Export a sheet as JSON (typed or string, with or without headers)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		rng, _ := cmd.Flags().GetString("range")
		withHeaders, _ := cmd.Flags().GetBool("with-headers")
		typed, _ := cmd.Flags().GetBool("typed")
		limit, _ := cmd.Flags().GetInt("limit")

		data, err := excel.ToJSON(path, excel.JSONReadOptions{
			Sheet: sheet, Range: rng, WithHeaders: withHeaders, Typed: typed, Limit: limit,
		})
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		output.PrintJSON(data)
		return nil
	},
}
