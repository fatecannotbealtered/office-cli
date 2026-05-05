package cmd

import (
	"fmt"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/excel"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// Excel sheet commands: rename-sheet, delete-sheet, copy-sheet,
// insert/delete rows & cols, hide/show sheet, validation, hyperlink
// ---------------------------------------------------------------------------

var excelRenameSheetCmd = &cobra.Command{
	Use:   "rename-sheet <FILE>",
	Short: "Rename a worksheet",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		from, _ := cmd.Flags().GetString("from")
		to, _ := cmd.Flags().GetString("to")
		if from == "" || to == "" {
			return emitError("--from and --to are required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("rename sheet", map[string]any{"file": path, "from": from, "to": to}) {
			return nil
		}

		if err := excel.RenameSheet(path, from, to); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "from": from, "to": to})
			return nil
		}
		output.Success(fmt.Sprintf("renamed %q -> %q", from, to))
		return nil
	},
}

var excelDeleteSheetCmd = &cobra.Command{
	Use:   "delete-sheet <FILE>",
	Short: "Delete a worksheet (last remaining sheet cannot be deleted)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		if sheet == "" {
			return emitError("--sheet is required", output.ErrValidation, "", ExitBadArgs)
		}

		if !forceMode && !confirmAction(fmt.Sprintf("Delete sheet %q? Type the sheet name to confirm", sheet), sheet) {
			return emitError("aborted", output.ErrValidation, path, ExitBadArgs)
		}

		if dryRunOutput("delete sheet", map[string]any{"file": path, "sheet": sheet}) {
			return nil
		}

		if err := excel.DeleteSheet(path, sheet); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheet": sheet})
			return nil
		}
		output.Success(fmt.Sprintf("deleted sheet %q from %s", sheet, path))
		return nil
	},
}

var excelCopySheetCmd = &cobra.Command{
	Use:   "copy-sheet <FILE>",
	Short: "Duplicate a worksheet (styles/formatting only; values not copied)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		from, _ := cmd.Flags().GetString("from")
		to, _ := cmd.Flags().GetString("to")
		if from == "" || to == "" {
			return emitError("--from and --to are required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("copy sheet", map[string]any{"file": path, "from": from, "to": to}) {
			return nil
		}

		idx, err := excel.CopySheet(path, from, to)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "from": from, "to": to, "index": idx})
			return nil
		}
		output.Success(fmt.Sprintf("copied %q -> %q (index %d)", from, to, idx))
		return nil
	},
}

var excelInsertRowsCmd = &cobra.Command{
	Use:   "insert-rows <FILE>",
	Short: "Insert blank rows after a given row",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		after, _ := cmd.Flags().GetInt("after")
		count, _ := cmd.Flags().GetInt("count")
		if after < 1 {
			return emitError("--after must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.InsertRows(path, sheet, after, count); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "after": after, "count": count})
			return nil
		}
		output.Success(fmt.Sprintf("inserted %d row(s) after row %d in %s", count, after, path))
		return nil
	},
}

var excelInsertColsCmd = &cobra.Command{
	Use:   "insert-cols <FILE>",
	Short: "Insert blank columns after a given column",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		after, _ := cmd.Flags().GetInt("after")
		count, _ := cmd.Flags().GetInt("count")
		if after < 1 {
			return emitError("--after must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.InsertCols(path, sheet, after, count); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "after": after, "count": count})
			return nil
		}
		output.Success(fmt.Sprintf("inserted %d column(s) after column %d in %s", count, after, path))
		return nil
	},
}

var excelDeleteRowsCmd = &cobra.Command{
	Use:   "delete-rows <FILE>",
	Short: "Delete rows starting at a given row",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		from, _ := cmd.Flags().GetInt("from")
		count, _ := cmd.Flags().GetInt("count")
		if from < 1 {
			return emitError("--from must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.DeleteRows(path, sheet, from, count); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "from": from, "count": count})
			return nil
		}
		output.Success(fmt.Sprintf("deleted %d row(s) starting at row %d in %s", count, from, path))
		return nil
	},
}

var excelDeleteColsCmd = &cobra.Command{
	Use:   "delete-cols <FILE>",
	Short: "Delete columns starting at a given column",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		from, _ := cmd.Flags().GetInt("from")
		count, _ := cmd.Flags().GetInt("count")
		if from < 1 {
			return emitError("--from must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.DeleteCols(path, sheet, from, count); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "from": from, "count": count})
			return nil
		}
		output.Success(fmt.Sprintf("deleted %d column(s) starting at column %d in %s", count, from, path))
		return nil
	},
}

var excelHideSheetCmd = &cobra.Command{
	Use:   "hide-sheet <FILE>",
	Short: "Hide a worksheet",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		if sheet == "" {
			return emitError("--sheet is required", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.SetSheetVisible(path, sheet, false); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheet": sheet, "visible": false})
			return nil
		}
		output.Success(fmt.Sprintf("hid sheet %q in %s", sheet, path))
		return nil
	},
}

var excelShowSheetCmd = &cobra.Command{
	Use:   "show-sheet <FILE>",
	Short: "Show (unhide) a worksheet",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		if sheet == "" {
			return emitError("--sheet is required", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.SetSheetVisible(path, sheet, true); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "sheet": sheet, "visible": true})
			return nil
		}
		output.Success(fmt.Sprintf("showed sheet %q in %s", sheet, path))
		return nil
	},
}

var excelValidationCmd = &cobra.Command{
	Use:   "validation <FILE>",
	Short: "Add data validation (dropdown lists, number/date ranges, custom formulas)",
	Long: `Add data validation rules to a worksheet. Supports dropdown lists,
numeric/date/time ranges with operators, and custom formulas.`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		vType, _ := cmd.Flags().GetString("type")
		if cellRange == "" {
			return emitError("--range is required (e.g. D2:D100)", output.ErrValidation, "", ExitBadArgs)
		}
		if vType == "" {
			return emitError("--type is required (list, whole, decimal, date, time, textLength, custom)", output.ErrValidation, "", ExitBadArgs)
		}

		listRaw, _ := cmd.Flags().GetString("list")
		var listItems []string
		if listRaw != "" {
			for _, item := range strings.Split(listRaw, ",") {
				if trimmed := strings.TrimSpace(item); trimmed != "" {
					listItems = append(listItems, trimmed)
				}
			}
		}

		spec := excel.ValidationSpec{
			Type:        vType,
			Range:       cellRange,
			Operator:    cmd.Flag("operator").Value.String(),
			Min:         cmd.Flag("min").Value.String(),
			Max:         cmd.Flag("max").Value.String(),
			List:        listItems,
			Formula:     cmd.Flag("formula").Value.String(),
			ErrorMsg:    cmd.Flag("error-msg").Value.String(),
			ErrorTitle:  cmd.Flag("error-title").Value.String(),
			PromptMsg:   cmd.Flag("prompt-msg").Value.String(),
			PromptTitle: cmd.Flag("prompt-title").Value.String(),
		}
		spec.AllowBlank, _ = cmd.Flags().GetBool("allow-blank")

		if dryRunOutput("add validation", map[string]any{"file": path, "range": cellRange, "type": vType}) {
			return nil
		}

		if err := excel.AddValidation(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange, "type": vType})
			return nil
		}
		output.Success(fmt.Sprintf("added %s validation to %s in %s", vType, cellRange, path))
		return nil
	},
}

var excelHyperlinkCmd = &cobra.Command{
	Use:   "hyperlink <FILE>",
	Short: "Set or get a cell hyperlink",
	Long: `Set a hyperlink on a cell, or read an existing one.

Set mode (default):
  office-cli excel hyperlink report.xlsx --cell A1 --link "https://example.com" --display "Click here"

Get mode (--get):
  office-cli excel hyperlink report.xlsx --sheet Data --cell A1 --get --json`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cell, _ := cmd.Flags().GetString("cell")
		if cell == "" {
			return emitError("--cell is required", output.ErrValidation, "", ExitBadArgs)
		}
		getMode, _ := cmd.Flags().GetBool("get")

		if getMode {
			result, err := excel.GetHyperlink(path, sheet, cell)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(result)
				return nil
			}
			if !result.Exists {
				output.Info(fmt.Sprintf("no hyperlink on %s", cell))
			} else {
				output.Info(fmt.Sprintf("%s → %s", cell, result.Link))
			}
			return nil
		}

		// Set mode.
		link, _ := cmd.Flags().GetString("link")
		if link == "" {
			return emitError("--link is required for set mode", output.ErrValidation, "", ExitBadArgs)
		}
		linkType, _ := cmd.Flags().GetString("type")
		display, _ := cmd.Flags().GetString("display")
		tooltip, _ := cmd.Flags().GetString("tooltip")

		spec := excel.HyperlinkSpec{
			Cell: cell, Link: link, Type: linkType, Display: display, Tooltip: tooltip,
		}

		if dryRunOutput("set hyperlink", map[string]any{"file": path, "cell": cell, "link": link}) {
			return nil
		}

		if err := excel.SetHyperlink(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "cell": cell, "link": link})
			return nil
		}
		output.Success(fmt.Sprintf("set hyperlink on %s → %s", cell, link))
		return nil
	},
}

func init() {
	excelRenameSheetCmd.Flags().String("from", "", "Existing sheet name (required)")
	excelRenameSheetCmd.Flags().String("to", "", "New sheet name (required)")
	excelCmd.AddCommand(excelRenameSheetCmd)

	excelDeleteSheetCmd.Flags().String("sheet", "", "Sheet name to delete (required)")
	excelCmd.AddCommand(excelDeleteSheetCmd)

	excelCopySheetCmd.Flags().String("from", "", "Source sheet name (required)")
	excelCopySheetCmd.Flags().String("to", "", "Target sheet name (required)")
	excelCmd.AddCommand(excelCopySheetCmd)

	excelInsertRowsCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelInsertRowsCmd.Flags().Int("after", 0, "Row number after which to insert (1-based)")
	excelInsertRowsCmd.Flags().Int("count", 1, "Number of rows to insert")
	excelCmd.AddCommand(excelInsertRowsCmd)

	excelInsertColsCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelInsertColsCmd.Flags().Int("after", 0, "Column number after which to insert (1-based)")
	excelInsertColsCmd.Flags().Int("count", 1, "Number of columns to insert")
	excelCmd.AddCommand(excelInsertColsCmd)

	excelDeleteRowsCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelDeleteRowsCmd.Flags().Int("from", 0, "First row to delete (1-based)")
	excelDeleteRowsCmd.Flags().Int("count", 1, "Number of rows to delete")
	excelCmd.AddCommand(excelDeleteRowsCmd)

	excelDeleteColsCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelDeleteColsCmd.Flags().Int("from", 0, "First column to delete (1-based)")
	excelDeleteColsCmd.Flags().Int("count", 1, "Number of columns to delete")
	excelCmd.AddCommand(excelDeleteColsCmd)

	excelHideSheetCmd.Flags().String("sheet", "", "Sheet name to hide (required)")
	excelCmd.AddCommand(excelHideSheetCmd)

	excelShowSheetCmd.Flags().String("sheet", "", "Sheet name to show (required)")
	excelCmd.AddCommand(excelShowSheetCmd)

	excelValidationCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelValidationCmd.Flags().String("range", "", "Cell range (e.g. D2:D100) — required")
	excelValidationCmd.Flags().String("type", "", "Validation type: list, whole, decimal, date, time, textLength, custom — required")
	excelValidationCmd.Flags().String("operator", "", "Operator: between, equal, greaterThan, lessThan, etc. (default: between)")
	excelValidationCmd.Flags().String("min", "", "Minimum value for range validation")
	excelValidationCmd.Flags().String("max", "", "Maximum value for range validation")
	excelValidationCmd.Flags().String("list", "", "Comma-separated list items for dropdown (e.g. \"A,B,C\")")
	excelValidationCmd.Flags().String("formula", "", "Custom validation formula (e.g. \"=LEN(E2)>0\")")
	excelValidationCmd.Flags().String("error-msg", "", "Error message shown on invalid input")
	excelValidationCmd.Flags().String("error-title", "", "Error dialog title")
	excelValidationCmd.Flags().String("prompt-msg", "", "Input prompt message")
	excelValidationCmd.Flags().String("prompt-title", "", "Input prompt title")
	excelValidationCmd.Flags().Bool("allow-blank", false, "Allow blank cells")
	excelCmd.AddCommand(excelValidationCmd)

	excelHyperlinkCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelHyperlinkCmd.Flags().String("cell", "", "Cell reference (e.g. A1) — required")
	excelHyperlinkCmd.Flags().String("link", "", "URL or 'Sheet1!A1' for internal — required for set")
	excelHyperlinkCmd.Flags().String("type", "External", "Link type: External or Location")
	excelHyperlinkCmd.Flags().String("display", "", "Display text (default: current cell value)")
	excelHyperlinkCmd.Flags().String("tooltip", "", "Hover tooltip")
	excelHyperlinkCmd.Flags().Bool("get", false, "Get mode: read the hyperlink for a cell")
	excelCmd.AddCommand(excelHyperlinkCmd)

	markWrite(excelRenameSheetCmd)
	markFull(excelDeleteSheetCmd)
	markWrite(excelCopySheetCmd)
	markWrite(excelInsertRowsCmd)
	markWrite(excelInsertColsCmd)
	markWrite(excelDeleteRowsCmd)
	markWrite(excelDeleteColsCmd)
	markWrite(excelHideSheetCmd)
	markWrite(excelShowSheetCmd)
	markWrite(excelValidationCmd)
	markWrite(excelHyperlinkCmd)
}
