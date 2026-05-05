package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/fatecannotbealtered/office-cli/internal/engine/word"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// Word edit commands: delete, insert-before/after, update-table-cell,
// update-table, search, add/delete table rows/cols, style, style-table
// ---------------------------------------------------------------------------

var wordDeleteCmd = &cobra.Command{
	Use:   "delete <FILE>",
	Short: "Delete a body element (paragraph or table) by index",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		index, _ := cmd.Flags().GetInt("index")
		out, _ := cmd.Flags().GetString("output")
		if index < 0 {
			return emitError("--index is required (0-based body element index)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("delete element", map[string]any{"file": path, "index": index}) {
			return nil
		}

		if err := word.DeleteBodyElement(path, out, index); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "deleted": index})
			return nil
		}
		output.Success(fmt.Sprintf("deleted element %d from %s", index, out))
		return nil
	},
}

var wordInsertBeforeCmd = &cobra.Command{
	Use:   "insert-before <FILE>",
	Short: "Insert a body element before the given index",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return wordInsertAt(cmd, args, true)
	},
}

var wordInsertAfterCmd = &cobra.Command{
	Use:   "insert-after <FILE>",
	Short: "Insert a body element after the given index",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		return wordInsertAt(cmd, args, false)
	},
}

func wordInsertAt(cmd *cobra.Command, args []string, before bool) error {
	path, err := resolveInput(args[0], "docx")
	if err != nil {
		return err
	}
	index, _ := cmd.Flags().GetInt("index")
	contentType, _ := cmd.Flags().GetString("type")
	text, _ := cmd.Flags().GetString("text")
	style, _ := cmd.Flags().GetString("style")
	level, _ := cmd.Flags().GetInt("level")
	rowsRaw, _ := cmd.Flags().GetString("rows")
	out, _ := cmd.Flags().GetString("output")

	if index < 0 {
		return emitError("--index is required (0-based body element index)", output.ErrValidation, "", ExitBadArgs)
	}
	if contentType == "" {
		return emitError("--type is required (paragraph | heading | table | page-break)", output.ErrValidation, "", ExitBadArgs)
	}
	if out == "" {
		out = path
	} else {
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}
	}

	// Build the XML content based on type
	var xmlContent string
	switch contentType {
	case "paragraph":
		if text == "" {
			return emitError("--text is required for --type paragraph", output.ErrValidation, "", ExitBadArgs)
		}
		xmlContent = word.ParagraphXML(word.Paragraph{Style: style, Text: text})
	case "heading":
		if text == "" {
			return emitError("--text is required for --type heading", output.ErrValidation, "", ExitBadArgs)
		}
		if level < 1 || level > 6 {
			return emitError("--level must be 1-6", output.ErrValidation, "", ExitBadArgs)
		}
		xmlContent = word.ParagraphXML(word.Paragraph{Style: fmt.Sprintf("Heading%d", level), Text: text})
	case "table":
		if rowsRaw == "" {
			return emitError("--rows is required for --type table", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(rowsRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var rows [][]string
		if err := json.Unmarshal(data, &rows); err != nil {
			return emitError("invalid --rows JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		xmlContent = word.TableXML(rows)
	case "page-break":
		xmlContent = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`
	default:
		return emitError("--type must be paragraph | heading | table | page-break", output.ErrValidation, "", ExitBadArgs)
	}

	action := "insert-before"
	if !before {
		action = "insert-after"
	}

	if dryRunOutput(action, map[string]any{"file": path, "index": index, "type": contentType}) {
		return nil
	}

	if err := word.InsertBodyElement(path, out, index, before, xmlContent); err != nil {
		return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
	}
	if jsonMode {
		output.PrintJSON(map[string]any{"status": "ok", "file": out, "index": index, "type": contentType})
		return nil
	}
	output.Success(fmt.Sprintf("%s %s at index %d in %s", action, contentType, index, out))
	return nil
}

var wordUpdateTableCellCmd = &cobra.Command{
	Use:   "update-table-cell <FILE>",
	Short: "Update a single cell in a table by body-element index, row, and column",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		row, _ := cmd.Flags().GetInt("row")
		col, _ := cmd.Flags().GetInt("col")
		value, _ := cmd.Flags().GetString("value")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		if row < 0 {
			return emitError("--row is required (0-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if col < 0 {
			return emitError("--col is required (0-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("update table cell", map[string]any{"file": path, "tableIndex": tableIndex, "row": row, "col": col}) {
			return nil
		}

		if err := word.UpdateTableCell(path, out, tableIndex, row, col, value); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex, "row": row, "col": col, "value": value})
			return nil
		}
		output.Success(fmt.Sprintf("updated table[%d].row[%d].col[%d] = %q in %s", tableIndex, row, col, value, out))
		return nil
	},
}

var wordSearchCmd = &cobra.Command{
	Use:   "search <FILE>",
	Short: "Search for a keyword across paragraphs and table cells",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		keyword, _ := cmd.Flags().GetString("keyword")
		if keyword == "" {
			return emitError("--keyword is required", output.ErrValidation, "", ExitBadArgs)
		}

		results, err := word.SearchBodyElements(path, keyword)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "keyword": keyword, "hits": len(results), "results": results})
			return nil
		}
		if len(results) == 0 {
			output.Gray(fmt.Sprintf("no matches for %q", keyword))
			return nil
		}
		rows := make([][]string, 0, len(results))
		for _, r := range results {
			if r.ElementType == "paragraph" {
				rows = append(rows, []string{
					fmt.Sprintf("%d", r.ElementIndex),
					"paragraph",
					r.Style,
					r.Text,
					"",
				})
			} else {
				rows = append(rows, []string{
					fmt.Sprintf("%d", r.ElementIndex),
					"table",
					fmt.Sprintf("[%d,%d]", r.Row, r.Col),
					r.Cell,
					"",
				})
			}
		}
		output.Table([]string{"#", "type", "location", "match", ""}, rows)
		return nil
	},
}

// ---------------------------------------------------------------------------
// Table row operations
// ---------------------------------------------------------------------------

var wordAddTableRowsCmd = &cobra.Command{
	Use:   "add-table-rows <FILE>",
	Short: "Add rows to an existing table in a .docx",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		position, _ := cmd.Flags().GetInt("position")
		rowsRaw, _ := cmd.Flags().GetString("rows")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		if rowsRaw == "" {
			return emitError("--rows is required (JSON array of arrays)", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(rowsRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var rows [][]string
		if err := json.Unmarshal(data, &rows); err != nil {
			return emitError("invalid --rows JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add table rows", map[string]any{"file": path, "tableIndex": tableIndex, "newRows": len(rows)}) {
			return nil
		}

		if err := word.AddTableRows(path, out, tableIndex, position, rows); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex, "added": len(rows)})
			return nil
		}
		output.Success(fmt.Sprintf("added %d row(s) to table %d in %s", len(rows), tableIndex, out))
		return nil
	},
}

var wordDeleteTableRowsCmd = &cobra.Command{
	Use:   "delete-table-rows <FILE>",
	Short: "Delete rows from a table in a .docx",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		startRow, _ := cmd.Flags().GetInt("start-row")
		endRow, _ := cmd.Flags().GetInt("end-row")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		if startRow < 0 {
			return emitError("--start-row is required (0-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if endRow < 0 {
			return emitError("--end-row is required (0-based, inclusive)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("delete table rows", map[string]any{"file": path, "tableIndex": tableIndex, "startRow": startRow, "endRow": endRow}) {
			return nil
		}

		if err := word.DeleteTableRows(path, out, tableIndex, startRow, endRow); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex, "deletedRows": endRow - startRow + 1})
			return nil
		}
		output.Success(fmt.Sprintf("deleted rows %d-%d from table %d in %s", startRow, endRow, tableIndex, out))
		return nil
	},
}

// ---------------------------------------------------------------------------
// Table column operations
// ---------------------------------------------------------------------------

var wordAddTableColsCmd = &cobra.Command{
	Use:   "add-table-cols <FILE>",
	Short: "Add columns to an existing table in a .docx",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		position, _ := cmd.Flags().GetInt("position")
		valuesRaw, _ := cmd.Flags().GetString("values")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		var values []string
		if valuesRaw != "" {
			data, err := readSpecArg(valuesRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &values); err != nil {
				return emitError("invalid --values JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add table columns", map[string]any{"file": path, "tableIndex": tableIndex}) {
			return nil
		}

		if err := word.AddTableCols(path, out, tableIndex, position, values); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex})
			return nil
		}
		output.Success(fmt.Sprintf("added column to table %d in %s", tableIndex, out))
		return nil
	},
}

var wordDeleteTableColsCmd = &cobra.Command{
	Use:   "delete-table-cols <FILE>",
	Short: "Delete columns from a table in a .docx",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		startCol, _ := cmd.Flags().GetInt("start-col")
		endCol, _ := cmd.Flags().GetInt("end-col")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		if startCol < 0 {
			return emitError("--start-col is required (0-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if endCol < 0 {
			return emitError("--end-col is required (0-based, inclusive)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("delete table columns", map[string]any{"file": path, "tableIndex": tableIndex, "startCol": startCol, "endCol": endCol}) {
			return nil
		}

		if err := word.DeleteTableCols(path, out, tableIndex, startCol, endCol); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex, "deletedCols": endCol - startCol + 1})
			return nil
		}
		output.Success(fmt.Sprintf("deleted columns %d-%d from table %d in %s", startCol, endCol, tableIndex, out))
		return nil
	},
}

// ---------------------------------------------------------------------------
// Batch table cell update
// ---------------------------------------------------------------------------

var wordUpdateTableCmd = &cobra.Command{
	Use:   "update-table <FILE>",
	Short: "Batch-update multiple cells in a table (--spec is a JSON array of {row,col,value})",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		specRaw, _ := cmd.Flags().GetString("spec")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}
		if specRaw == "" {
			return emitError("--spec is required (JSON array of {row,col,value})", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var updates []word.CellUpdate
		if err := json.Unmarshal(data, &updates); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("batch update table cells", map[string]any{"file": path, "tableIndex": tableIndex, "updates": len(updates)}) {
			return nil
		}

		if err := word.UpdateTableCells(path, out, tableIndex, updates); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "tableIndex": tableIndex, "updated": len(updates)})
			return nil
		}
		output.Success(fmt.Sprintf("updated %d cell(s) in table %d -> %s", len(updates), tableIndex, out))
		return nil
	},
}

// ---------------------------------------------------------------------------
// word style — paragraph styling
// ---------------------------------------------------------------------------

var wordStyleCmd = &cobra.Command{
	Use:   "style <FILE>",
	Short: "Apply text and paragraph formatting to a body element (paragraph or table)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		index, _ := cmd.Flags().GetInt("index")
		out, _ := cmd.Flags().GetString("output")

		// Run style flags
		bold, _ := cmd.Flags().GetBool("bold")
		italic, _ := cmd.Flags().GetBool("italic")
		underline, _ := cmd.Flags().GetBool("underline")
		strike, _ := cmd.Flags().GetBool("strikethrough")
		fontFamily, _ := cmd.Flags().GetString("font-family")
		fontSize, _ := cmd.Flags().GetInt("font-size")
		color, _ := cmd.Flags().GetString("color")

		// Paragraph style flags
		align, _ := cmd.Flags().GetString("align")
		spaceBefore, _ := cmd.Flags().GetInt("space-before")
		spaceAfter, _ := cmd.Flags().GetInt("space-after")
		lineSpacing, _ := cmd.Flags().GetInt("line-spacing")
		indentLeft, _ := cmd.Flags().GetInt("indent-left")
		indentRight, _ := cmd.Flags().GetInt("indent-right")
		firstLine, _ := cmd.Flags().GetInt("first-line")

		if index < 0 {
			return emitError("--index is required", output.ErrValidation, "", ExitBadArgs)
		}

		runStyle := word.RunStyle{
			Bold:       bold,
			Italic:     italic,
			Underline:  underline,
			Strike:     strike,
			FontFamily: fontFamily,
			FontSize:   fontSize,
			Color:      color,
		}
		paraStyle := word.ParaStyle{
			Align:       align,
			SpaceBefore: spaceBefore,
			SpaceAfter:  spaceAfter,
			LineSpacing: lineSpacing,
			IndentLeft:  indentLeft,
			IndentRight: indentRight,
			FirstLine:   firstLine,
		}

		if runStyle.IsZero() && paraStyle.IsZero() {
			return emitError("at least one style flag is required", output.ErrValidation, "", ExitBadArgs)
		}

		outPath, err := resolveOutputOrDefault(path, out)
		if err != nil {
			return err
		}

		if dryRunOutput("style element", map[string]any{"file": path, "index": index, "runStyle": runStyle, "paraStyle": paraStyle}) {
			return nil
		}

		if err := word.StyleParagraph(path, outPath, index, runStyle, paraStyle); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": outPath, "index": index})
			return nil
		}
		output.Success(fmt.Sprintf("styled element %d -> %s", index, outPath))
		return nil
	},
}

// ---------------------------------------------------------------------------
// word style-table — table cell styling
// ---------------------------------------------------------------------------

var wordStyleTableCmd = &cobra.Command{
	Use:   "style-table <FILE>",
	Short: "Apply cell and text formatting to a range of cells in a table",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		tableIndex, _ := cmd.Flags().GetInt("table-index")
		startRow, _ := cmd.Flags().GetInt("start-row")
		endRow, _ := cmd.Flags().GetInt("end-row")
		startCol, _ := cmd.Flags().GetInt("start-col")
		endCol, _ := cmd.Flags().GetInt("end-col")
		specRaw, _ := cmd.Flags().GetString("spec")
		out, _ := cmd.Flags().GetString("output")

		if tableIndex < 0 {
			return emitError("--table-index is required", output.ErrValidation, "", ExitBadArgs)
		}

		outPath, err := resolveOutputOrDefault(path, out)
		if err != nil {
			return err
		}

		// Batch spec mode
		if specRaw != "" {
			return styleTableBatch(path, outPath, tableIndex, specRaw)
		}

		// Individual flag mode
		bgColor, _ := cmd.Flags().GetString("bg-color")
		borders, _ := cmd.Flags().GetString("borders")
		vAlign, _ := cmd.Flags().GetString("valign")

		bold, _ := cmd.Flags().GetBool("bold")
		italic, _ := cmd.Flags().GetBool("italic")
		underline, _ := cmd.Flags().GetBool("underline")
		strike, _ := cmd.Flags().GetBool("strikethrough")
		fontFamily, _ := cmd.Flags().GetString("font-family")
		fontSize, _ := cmd.Flags().GetInt("font-size")
		color, _ := cmd.Flags().GetString("color")

		cellStyle := word.CellStyle{
			BgColor: bgColor,
			Borders: borders,
			VAlign:  vAlign,
		}
		runStyle := word.RunStyle{
			Bold:       bold,
			Italic:     italic,
			Underline:  underline,
			Strike:     strike,
			FontFamily: fontFamily,
			FontSize:   fontSize,
			Color:      color,
		}

		if cellStyle.IsZero() && runStyle.IsZero() {
			return emitError("at least one style flag is required (or use --spec for batch)", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("style table cells", map[string]any{"file": path, "tableIndex": tableIndex, "cellStyle": cellStyle, "runStyle": runStyle}) {
			return nil
		}

		if err := word.StyleTable(path, outPath, tableIndex, startRow, endRow, startCol, endCol, cellStyle, runStyle); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": outPath, "tableIndex": tableIndex})
			return nil
		}
		output.Success(fmt.Sprintf("styled table %d -> %s", tableIndex, outPath))
		return nil
	},
}

// styleTableSpec is one entry in the --spec JSON array for style-table batch mode.
type styleTableSpec struct {
	StartRow int    `json:"startRow"`
	EndRow   int    `json:"endRow"`
	StartCol int    `json:"startCol"`
	EndCol   int    `json:"endCol"`
	BgColor  string `json:"bgColor"`
	Borders  string `json:"borders"`
	VAlign   string `json:"vAlign"`
	Bold     bool   `json:"bold"`
	Italic   bool   `json:"italic"`
	Strike   bool   `json:"strike"`
	Color    string `json:"color"`
	FontSize int    `json:"fontSize"`
}

func styleTableBatch(path, outPath string, tableIndex int, specRaw string) error {
	data, err := readSpecArg(specRaw)
	if err != nil {
		return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
	}
	var specs []styleTableSpec
	if err := json.Unmarshal(data, &specs); err != nil {
		return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
	}

	if dryRunOutput("batch style table", map[string]any{"file": path, "tableIndex": tableIndex, "specs": len(specs)}) {
		return nil
	}

	for _, sp := range specs {
		cs := word.CellStyle{BgColor: sp.BgColor, Borders: sp.Borders, VAlign: sp.VAlign}
		rs := word.RunStyle{Bold: sp.Bold, Italic: sp.Italic, Strike: sp.Strike, Color: sp.Color, FontSize: sp.FontSize}
		sr, er := sp.StartRow, sp.EndRow
		sc, ec := sp.StartCol, sp.EndCol
		if err := word.StyleTable(path, outPath, tableIndex, sr, er, sc, ec, cs, rs); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		// Subsequent passes read from outPath
		path = outPath
	}
	if jsonMode {
		output.PrintJSON(map[string]any{"status": "ok", "file": outPath, "tableIndex": tableIndex, "styled": len(specs)})
		return nil
	}
	output.Success(fmt.Sprintf("styled %d region(s) in table %d -> %s", len(specs), tableIndex, outPath))
	return nil
}

// resolveOutputOrDefault returns out if non-empty, otherwise returns path.
func resolveOutputOrDefault(path, out string) (string, error) {
	if out == "" {
		return path, nil
	}
	return resolveOutput(out)
}

func init() {
	wordDeleteCmd.Flags().Int("index", -1, "0-based body element index to delete (required)")
	wordDeleteCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordDeleteCmd)

	wordInsertBeforeCmd.Flags().Int("index", -1, "0-based body element index to insert before (required)")
	wordInsertBeforeCmd.Flags().String("type", "", "Content type: paragraph | heading | table | page-break (required)")
	wordInsertBeforeCmd.Flags().String("text", "", "Text for paragraph/heading")
	wordInsertBeforeCmd.Flags().String("style", "", "Style for paragraph (Normal, Heading1-6, Title)")
	wordInsertBeforeCmd.Flags().Int("level", 1, "Heading level 1-6 (for --type heading)")
	wordInsertBeforeCmd.Flags().String("rows", "", "JSON array of arrays for table (for --type table)")
	wordInsertBeforeCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordInsertBeforeCmd)

	wordInsertAfterCmd.Flags().Int("index", -1, "0-based body element index to insert after (required)")
	wordInsertAfterCmd.Flags().String("type", "", "Content type: paragraph | heading | table | page-break (required)")
	wordInsertAfterCmd.Flags().String("text", "", "Text for paragraph/heading")
	wordInsertAfterCmd.Flags().String("style", "", "Style for paragraph (Normal, Heading1-6, Title)")
	wordInsertAfterCmd.Flags().Int("level", 1, "Heading level 1-6 (for --type heading)")
	wordInsertAfterCmd.Flags().String("rows", "", "JSON array of arrays for table (for --type table)")
	wordInsertAfterCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordInsertAfterCmd)

	wordUpdateTableCellCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordUpdateTableCellCmd.Flags().Int("row", -1, "0-based row index (required)")
	wordUpdateTableCellCmd.Flags().Int("col", -1, "0-based column index (required)")
	wordUpdateTableCellCmd.Flags().String("value", "", "New cell value (required)")
	wordUpdateTableCellCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordUpdateTableCellCmd)

	wordSearchCmd.Flags().String("keyword", "", "Search keyword (required, case-insensitive)")
	wordCmd.AddCommand(wordSearchCmd)

	wordAddTableRowsCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordAddTableRowsCmd.Flags().Int("position", -1, "0-based row index to insert before (-1 = append)")
	wordAddTableRowsCmd.Flags().String("rows", "", "JSON array of arrays for new rows (required)")
	wordAddTableRowsCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddTableRowsCmd)

	wordDeleteTableRowsCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordDeleteTableRowsCmd.Flags().Int("start-row", -1, "0-based start row index (required)")
	wordDeleteTableRowsCmd.Flags().Int("end-row", -1, "0-based end row index, inclusive (required)")
	wordDeleteTableRowsCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordDeleteTableRowsCmd)

	wordAddTableColsCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordAddTableColsCmd.Flags().Int("position", -1, "0-based col index to insert before (-1 = append)")
	wordAddTableColsCmd.Flags().String("values", "", "JSON array of default values for the new column (one per row)")
	wordAddTableColsCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddTableColsCmd)

	wordDeleteTableColsCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordDeleteTableColsCmd.Flags().Int("start-col", -1, "0-based start col index (required)")
	wordDeleteTableColsCmd.Flags().Int("end-col", -1, "0-based end col index, inclusive (required)")
	wordDeleteTableColsCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordDeleteTableColsCmd)

	wordUpdateTableCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordUpdateTableCmd.Flags().String("spec", "", "JSON array of {row,col,value} objects (required)")
	wordUpdateTableCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordUpdateTableCmd)

	// --- Style commands ---
	wordStyleCmd.Flags().Int("index", -1, "0-based body element index (required)")
	wordStyleCmd.Flags().Bool("bold", false, "Bold text")
	wordStyleCmd.Flags().Bool("italic", false, "Italic text")
	wordStyleCmd.Flags().Bool("underline", false, "Underline text")
	wordStyleCmd.Flags().Bool("strikethrough", false, "Strikethrough text")
	wordStyleCmd.Flags().String("font-family", "", "Font family name")
	wordStyleCmd.Flags().Int("font-size", 0, "Font size in points")
	wordStyleCmd.Flags().String("color", "", "Text color hex (e.g. FF0000)")
	wordStyleCmd.Flags().String("align", "", "Paragraph alignment: left, center, right, both")
	wordStyleCmd.Flags().Int("space-before", 0, "Space before paragraph (twips)")
	wordStyleCmd.Flags().Int("space-after", 0, "Space after paragraph (twips)")
	wordStyleCmd.Flags().Int("line-spacing", 0, "Line spacing (twips, 240=single)")
	wordStyleCmd.Flags().Int("indent-left", 0, "Left indentation (twips)")
	wordStyleCmd.Flags().Int("indent-right", 0, "Right indentation (twips)")
	wordStyleCmd.Flags().Int("first-line", 0, "First line indent (twips)")
	wordStyleCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordStyleCmd)

	wordStyleTableCmd.Flags().Int("table-index", -1, "Body element index of the table (required)")
	wordStyleTableCmd.Flags().Int("start-row", -1, "0-based start row index (-1 = first)")
	wordStyleTableCmd.Flags().Int("end-row", -1, "0-based end row index, inclusive (-1 = last)")
	wordStyleTableCmd.Flags().Int("start-col", -1, "0-based start col index (-1 = first)")
	wordStyleTableCmd.Flags().Int("end-col", -1, "0-based end col index, inclusive (-1 = last)")
	wordStyleTableCmd.Flags().String("bg-color", "", "Background color hex (e.g. FFFF00)")
	wordStyleTableCmd.Flags().String("borders", "", "Border style: none, single, thick")
	wordStyleTableCmd.Flags().String("valign", "", "Vertical alignment: top, center, bottom")
	wordStyleTableCmd.Flags().Bool("bold", false, "Bold text")
	wordStyleTableCmd.Flags().Bool("italic", false, "Italic text")
	wordStyleTableCmd.Flags().Bool("underline", false, "Underline text")
	wordStyleTableCmd.Flags().Bool("strikethrough", false, "Strikethrough text")
	wordStyleTableCmd.Flags().String("font-family", "", "Font family name")
	wordStyleTableCmd.Flags().Int("font-size", 0, "Font size in points")
	wordStyleTableCmd.Flags().String("color", "", "Text color hex (e.g. FF0000)")
	wordStyleTableCmd.Flags().String("spec", "", "JSON array of style specs for batch mode")
	wordStyleTableCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordStyleTableCmd)

	markWrite(wordDeleteCmd)
	markWrite(wordInsertBeforeCmd)
	markWrite(wordInsertAfterCmd)
	markWrite(wordUpdateTableCellCmd)
	markWrite(wordAddTableRowsCmd)
	markWrite(wordDeleteTableRowsCmd)
	markWrite(wordAddTableColsCmd)
	markWrite(wordDeleteTableColsCmd)
	markWrite(wordUpdateTableCmd)
	markWrite(wordStyleCmd)
	markWrite(wordStyleTableCmd)
}
