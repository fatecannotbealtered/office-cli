package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/fatecannotbealtered/office-cli/internal/engine/excel"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// Excel format commands: style, sort, freeze, merge/unmerge, col-width,
// row-height, chart, multi-chart, add-image, cond-format, auto-filter,
// data-bar, icon-set, color-scale, default-font, cell-style
// ---------------------------------------------------------------------------

var excelStyleCmd = &cobra.Command{
	Use:   "style <FILE>",
	Short: "Apply cell formatting (bold, color, borders, alignment, number format)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")

		// Batch mode: --spec
		specRaw, _ := cmd.Flags().GetString("spec")
		if specRaw != "" {
			data, err := readSpecArg(specRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			var entries []excel.BatchStyleEntry
			if err := json.Unmarshal(data, &entries); err != nil {
				return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if dryRunOutput("batch style", map[string]any{"file": path, "entries": len(entries)}) {
				return nil
			}
			if err := excel.StyleCellsBatch(path, sheet, entries); err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"status": "ok", "file": path, "entries": len(entries)})
				return nil
			}
			output.Success(fmt.Sprintf("applied %d style entries to %s", len(entries), path))
			return nil
		}

		// Single-range mode (original behavior).
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required (e.g. A1:C5), or use --spec for batch mode", output.ErrValidation, "", ExitBadArgs)
		}
		bold, _ := cmd.Flags().GetBool("bold")
		italic, _ := cmd.Flags().GetBool("italic")
		underline, _ := cmd.Flags().GetBool("underline")
		strike, _ := cmd.Flags().GetBool("strike")
		fontSize, _ := cmd.Flags().GetFloat64("font-size")
		fontFamily, _ := cmd.Flags().GetString("font-family")
		bgColor, _ := cmd.Flags().GetString("bg-color")
		textColor, _ := cmd.Flags().GetString("text-color")
		numberFmt, _ := cmd.Flags().GetString("number-format")
		align, _ := cmd.Flags().GetString("align")
		valign, _ := cmd.Flags().GetString("valign")
		wrapText, _ := cmd.Flags().GetBool("wrap-text")
		textRotation, _ := cmd.Flags().GetInt("text-rotation")
		indent, _ := cmd.Flags().GetInt("indent")
		shrinkToFit, _ := cmd.Flags().GetBool("shrink-to-fit")

		styleSpec := excel.StyleSpec{
			Bold: bold, Italic: italic, Underline: underline, Strike: strike,
			FontSize: fontSize, FontFamily: fontFamily,
			BgColor: bgColor, TextColor: textColor,
			NumberFormat: numberFmt, Align: align, Valign: valign,
			WrapText: wrapText, TextRotation: textRotation, Indent: indent, ShrinkToFit: shrinkToFit,
		}

		if dryRunOutput("apply style", map[string]any{"file": path, "range": cellRange}) {
			return nil
		}

		if err := excel.StyleCells(path, sheet, cellRange, styleSpec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("applied style to %s in %s", cellRange, path))
		return nil
	},
}

var excelSortCmd = &cobra.Command{
	Use:   "sort <FILE>",
	Short: "Sort rows in a range by a column",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		byCol, _ := cmd.Flags().GetInt("by-col")
		ascending, _ := cmd.Flags().GetBool("ascending")
		if cellRange == "" {
			return emitError("--range is required (e.g. A1:D10)", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.Sort(path, sheet, cellRange, byCol, ascending); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange, "byCol": byCol, "ascending": ascending})
			return nil
		}
		dir := "ascending"
		if !ascending {
			dir = "descending"
		}
		output.Success(fmt.Sprintf("sorted %s by column %d (%s) in %s", cellRange, byCol, dir, path))
		return nil
	},
}

var excelFreezeCmd = &cobra.Command{
	Use:   "freeze <FILE>",
	Short: "Freeze panes at a cell position (e.g. A2 freezes top row)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cell, _ := cmd.Flags().GetString("cell")
		if cell == "" {
			return emitError("--cell is required (e.g. A2)", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.Freeze(path, sheet, cell); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "cell": cell})
			return nil
		}
		output.Success(fmt.Sprintf("froze panes at %s in %s", cell, path))
		return nil
	},
}

var excelMergeCmd = &cobra.Command{
	Use:   "merge <FILE>",
	Short: "Merge a range of cells into one",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required (e.g. A1:C3)", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.MergeCells(path, sheet, cellRange); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("merged cells %s in %s", cellRange, path))
		return nil
	},
}

var excelUnmergeCmd = &cobra.Command{
	Use:   "unmerge <FILE>",
	Short: "Unmerge a previously merged range of cells",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required (e.g. A1:C3)", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.UnmergeCells(path, sheet, cellRange); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("unmerged cells %s in %s", cellRange, path))
		return nil
	},
}

var excelSetColWidthCmd = &cobra.Command{
	Use:   "set-col-width <FILE>",
	Short: "Set the width of a column",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		col, _ := cmd.Flags().GetInt("col")
		width, _ := cmd.Flags().GetFloat64("width")
		if col < 1 {
			return emitError("--col must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if width <= 0 {
			return emitError("--width must be > 0", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.SetColWidth(path, sheet, col, width); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "col": col, "width": width})
			return nil
		}
		output.Success(fmt.Sprintf("set column %d width to %.1f in %s", col, width, path))
		return nil
	},
}

var excelSetRowHeightCmd = &cobra.Command{
	Use:   "set-row-height <FILE>",
	Short: "Set the height of a row",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		row, _ := cmd.Flags().GetInt("row")
		height, _ := cmd.Flags().GetFloat64("height")
		if row < 1 {
			return emitError("--row must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if height <= 0 {
			return emitError("--height must be > 0", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.SetRowHeight(path, sheet, row, height); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "row": row, "height": height})
			return nil
		}
		output.Success(fmt.Sprintf("set row %d height to %.1f in %s", row, height, path))
		return nil
	},
}

var excelChartCmd = &cobra.Command{
	Use:   "chart <FILE>",
	Short: "Create a chart on a worksheet",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cell, _ := cmd.Flags().GetString("cell")
		chartType, _ := cmd.Flags().GetString("type")
		dataRange, _ := cmd.Flags().GetString("data-range")
		title, _ := cmd.Flags().GetString("title")
		width, _ := cmd.Flags().GetInt("width")
		height, _ := cmd.Flags().GetInt("height")
		if cell == "" {
			return emitError("--cell is required (anchor cell for chart)", output.ErrValidation, "", ExitBadArgs)
		}
		if dataRange == "" {
			return emitError("--data-range is required (e.g. A1:B5)", output.ErrValidation, "", ExitBadArgs)
		}

		spec := excel.ChartSpec{
			Type: chartType, Cell: cell, DataRange: dataRange,
			Title: title, Width: width, Height: height,
		}

		if dryRunOutput("create chart", map[string]any{"file": path, "type": chartType, "dataRange": dataRange}) {
			return nil
		}

		if err := excel.AddChart(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "type": chartType, "cell": cell})
			return nil
		}
		output.Success(fmt.Sprintf("created %s chart anchored at %s in %s", chartType, cell, path))
		return nil
	},
}

var excelAddImageCmd = &cobra.Command{
	Use:   "add-image <FILE>",
	Short: "Insert an image anchored at a cell",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cell, _ := cmd.Flags().GetString("cell")
		imagePath, _ := cmd.Flags().GetString("image")
		if cell == "" {
			return emitError("--cell is required (anchor cell)", output.ErrValidation, "", ExitBadArgs)
		}
		if imagePath == "" {
			return emitError("--image is required (path to image file)", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("add image", map[string]any{"file": path, "cell": cell, "image": imagePath}) {
			return nil
		}

		if err := excel.AddImage(path, sheet, cell, imagePath); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "cell": cell, "image": imagePath})
			return nil
		}
		output.Success(fmt.Sprintf("added image at %s in %s", cell, path))
		return nil
	},
}

var excelCondFormatCmd = &cobra.Command{
	Use:   "cond-format <FILE>",
	Short: "Apply conditional formatting to a range",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		condType, _ := cmd.Flags().GetString("type")
		criteria, _ := cmd.Flags().GetString("criteria")
		value, _ := cmd.Flags().GetString("value")
		bgColor, _ := cmd.Flags().GetString("bg-color")
		textColor, _ := cmd.Flags().GetString("text-color")
		if cellRange == "" {
			return emitError("--range is required", output.ErrValidation, "", ExitBadArgs)
		}
		if condType == "" {
			return emitError("--type is required (cell, top, bottom, average, duplicate, unique, formula)", output.ErrValidation, "", ExitBadArgs)
		}

		spec := excel.CondSpec{
			Type: condType, Criteria: criteria, Value: value,
			BgColor: bgColor, TextColor: textColor,
		}

		if dryRunOutput("conditional format", map[string]any{"file": path, "range": cellRange, "type": condType}) {
			return nil
		}

		if err := excel.ConditionalFormat(path, sheet, cellRange, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange, "type": condType})
			return nil
		}
		output.Success(fmt.Sprintf("applied %s conditional format to %s in %s", condType, cellRange, path))
		return nil
	},
}

var excelAutoFilterCmd = &cobra.Command{
	Use:   "auto-filter <FILE>",
	Short: "Set auto-filter on a header row",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required (e.g. A1:D1)", output.ErrValidation, "", ExitBadArgs)
		}
		if err := excel.SetAutoFilter(path, sheet, cellRange); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("set auto-filter on %s in %s", cellRange, path))
		return nil
	},
}

var excelMultiChartCmd = &cobra.Command{
	Use:   "multi-chart <FILE>",
	Short: "Create a chart with multiple data series",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		specRaw, _ := cmd.Flags().GetString("spec")
		if specRaw == "" {
			return emitError("--spec is required (JSON object, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}

		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var spec excel.MultiChartSpec
		if err := json.Unmarshal(data, &spec); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("create multi-chart", map[string]any{"file": path, "type": spec.Type, "series": len(spec.Series)}) {
			return nil
		}

		if err := excel.AddMultiChart(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "type": spec.Type, "series": len(spec.Series)})
			return nil
		}
		output.Success(fmt.Sprintf("created %s chart with %d series in %s", spec.Type, len(spec.Series), path))
		return nil
	},
}

var excelDataBarCmd = &cobra.Command{
	Use:   "data-bar <FILE>",
	Short: "Apply data bar conditional formatting to a range",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required", output.ErrValidation, "", ExitBadArgs)
		}
		minColor, _ := cmd.Flags().GetString("min-color")
		maxColor, _ := cmd.Flags().GetString("max-color")
		barOnly, _ := cmd.Flags().GetBool("bar-only")

		spec := excel.DataBarSpec{
			Range:    cellRange,
			MinColor: minColor,
			MaxColor: maxColor,
		}
		if barOnly {
			f := false
			spec.ShowValue = &f
		}

		if dryRunOutput("data bar", map[string]any{"file": path, "range": cellRange}) {
			return nil
		}

		if err := excel.AddDataBar(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange})
			return nil
		}
		output.Success(fmt.Sprintf("applied data bar to %s in %s", cellRange, path))
		return nil
	},
}

var excelIconSetCmd = &cobra.Command{
	Use:   "icon-set <FILE>",
	Short: "Apply icon set conditional formatting to a range",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required", output.ErrValidation, "", ExitBadArgs)
		}
		style, _ := cmd.Flags().GetString("style")
		reverse, _ := cmd.Flags().GetBool("reverse")

		spec := excel.IconSetSpec{
			Range:   cellRange,
			Style:   style,
			Reverse: reverse,
		}

		if dryRunOutput("icon set", map[string]any{"file": path, "range": cellRange, "style": style}) {
			return nil
		}

		if err := excel.AddIconSet(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange, "style": style})
			return nil
		}
		output.Success(fmt.Sprintf("applied %s icon set to %s in %s", style, cellRange, path))
		return nil
	},
}

var excelColorScaleCmd = &cobra.Command{
	Use:   "color-scale <FILE>",
	Short: "Apply color scale (heatmap) conditional formatting to a range",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "xlsx")
		if err != nil {
			return err
		}
		sheet, _ := cmd.Flags().GetString("sheet")
		cellRange, _ := cmd.Flags().GetString("range")
		if cellRange == "" {
			return emitError("--range is required", output.ErrValidation, "", ExitBadArgs)
		}
		minColor, _ := cmd.Flags().GetString("min-color")
		maxColor, _ := cmd.Flags().GetString("max-color")
		midColor, _ := cmd.Flags().GetString("mid-color")
		minType, _ := cmd.Flags().GetString("min-type")
		maxType, _ := cmd.Flags().GetString("max-type")
		minValue, _ := cmd.Flags().GetString("min-value")
		maxValue, _ := cmd.Flags().GetString("max-value")
		midValue, _ := cmd.Flags().GetString("mid-value")

		spec := excel.ColorScaleSpec{
			Range: cellRange, MinColor: minColor, MaxColor: maxColor, MidColor: midColor,
			MinType: minType, MaxType: maxType, MinValue: minValue, MaxValue: maxValue, MidValue: midValue,
		}

		scaleType := "2-color"
		if midColor != "" {
			scaleType = "3-color"
		}

		if dryRunOutput("color scale", map[string]any{"file": path, "range": cellRange, "type": scaleType}) {
			return nil
		}

		if err := excel.AddColorScale(path, sheet, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "range": cellRange, "type": scaleType})
			return nil
		}
		output.Success(fmt.Sprintf("applied %s scale to %s in %s", scaleType, cellRange, path))
		return nil
	},
}

var excelDefaultFontCmd = &cobra.Command{
	Use:   "default-font <file>",
	Short: "Get or set the workbook's default font",
	Long:  `Read the current default font, or set a new default font for the workbook.`,
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		setFont, _ := cmd.Flags().GetString("set")
		if setFont == "" {
			// Get mode
			name, err := excel.GetDefaultFont(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "defaultFont": name})
				return nil
			}
			fmt.Printf("Default font: %s\n", name)
			return nil
		}

		// Set mode
		if err := excel.SetDefaultFont(path, setFont); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "defaultFont": setFont})
			return nil
		}
		output.Success(fmt.Sprintf("default font set to %q", setFont))
		return nil
	},
}

var excelCellStyleCmd = &cobra.Command{
	Use:   "cell-style <file> <cell>",
	Short: "Read the style/font properties of a cell",
	Long:  `Returns font, alignment, fill, and number format properties for a single cell.`,
	Args:  cobra.ExactArgs(2),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		cell := args[1]
		sheet, _ := cmd.Flags().GetString("sheet")

		info, err := excel.GetCellStyleInfo(path, sheet, cell)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(info)
			return nil
		}
		// Human mode: pretty-print
		fmt.Printf("Cell: %s (style index: %d)\n", info.Cell, info.StyleIndex)
		if info.FontFamily != "" || info.FontSize > 0 || info.Bold || info.Italic {
			fmt.Printf("  Font: %s %gpt", info.FontFamily, info.FontSize)
			if info.Bold {
				fmt.Printf(" bold")
			}
			if info.Italic {
				fmt.Printf(" italic")
			}
			if info.Underline {
				fmt.Printf(" underline")
			}
			if info.Strike {
				fmt.Printf(" strike")
			}
			fmt.Println()
		}
		if info.FontColor != "" {
			fmt.Printf("  Font color: #%s\n", info.FontColor)
		}
		if info.BgColor != "" {
			fmt.Printf("  Background: #%s\n", info.BgColor)
		}
		if info.NumberFormat != "" {
			fmt.Printf("  Number format: %s\n", info.NumberFormat)
		}
		if info.Align != "" || info.Valign != "" {
			fmt.Printf("  Alignment: h=%s v=%s\n", info.Align, info.Valign)
		}
		return nil
	},
}

func init() {
	excelStyleCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelStyleCmd.Flags().String("range", "", "Cell range e.g. A1:C5 (required)")
	excelStyleCmd.Flags().Bool("bold", false, "Bold font")
	excelStyleCmd.Flags().Bool("italic", false, "Italic font")
	excelStyleCmd.Flags().Bool("underline", false, "Underline font")
	excelStyleCmd.Flags().Float64("font-size", 0, "Font size in points")
	excelStyleCmd.Flags().String("font-family", "", "Font family (e.g. Arial)")
	excelStyleCmd.Flags().String("bg-color", "", "Background color hex (e.g. FFFF00)")
	excelStyleCmd.Flags().String("text-color", "", "Text color hex (e.g. FF0000)")
	excelStyleCmd.Flags().String("number-format", "", "Number format (e.g. #,##0.00)")
	excelStyleCmd.Flags().String("align", "", "Horizontal alignment: left, center, right")
	excelStyleCmd.Flags().String("valign", "", "Vertical alignment: top, center, bottom")
	excelStyleCmd.Flags().Bool("wrap-text", false, "Wrap text")
	excelStyleCmd.Flags().Bool("strike", false, "Strikethrough")
	excelStyleCmd.Flags().Int("text-rotation", 0, "Text rotation (0-90 degrees, or 255 for vertical)")
	excelStyleCmd.Flags().Int("indent", 0, "Indent level (each level ≈ 3 spaces)")
	excelStyleCmd.Flags().Bool("shrink-to-fit", false, "Shrink font to fit cell width")
	excelStyleCmd.Flags().String("spec", "", "Batch style spec: JSON array of {range,style} objects (or @file.json)")
	excelCmd.AddCommand(excelStyleCmd)

	excelSortCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelSortCmd.Flags().String("range", "", "Cell range to sort (e.g. A1:D10) (required)")
	excelSortCmd.Flags().Int("by-col", 1, "Column number within the range to sort by (1-based)")
	excelSortCmd.Flags().Bool("ascending", true, "Sort ascending (false = descending)")
	excelCmd.AddCommand(excelSortCmd)

	excelFreezeCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelFreezeCmd.Flags().String("cell", "", "Freeze panes at this cell (e.g. A2 freezes top row) (required)")
	excelCmd.AddCommand(excelFreezeCmd)

	excelMergeCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelMergeCmd.Flags().String("range", "", "Cell range to merge (e.g. A1:C3) (required)")
	excelCmd.AddCommand(excelMergeCmd)

	excelUnmergeCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelUnmergeCmd.Flags().String("range", "", "Cell range to unmerge (e.g. A1:C3) (required)")
	excelCmd.AddCommand(excelUnmergeCmd)

	excelSetColWidthCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelSetColWidthCmd.Flags().Int("col", 0, "Column number (1-based) (required)")
	excelSetColWidthCmd.Flags().Float64("width", 0, "Column width (required)")
	excelCmd.AddCommand(excelSetColWidthCmd)

	excelSetRowHeightCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelSetRowHeightCmd.Flags().Int("row", 0, "Row number (1-based) (required)")
	excelSetRowHeightCmd.Flags().Float64("height", 0, "Row height (required)")
	excelCmd.AddCommand(excelSetRowHeightCmd)

	excelChartCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelChartCmd.Flags().String("cell", "", "Anchor cell for the chart (e.g. E2) (required)")
	excelChartCmd.Flags().String("type", "col", "Chart type: col, bar, line, area, pie, scatter")
	excelChartCmd.Flags().String("data-range", "", "Data range (e.g. A1:B5) (required)")
	excelChartCmd.Flags().String("title", "", "Chart title")
	excelChartCmd.Flags().Int("width", 12, "Chart width")
	excelChartCmd.Flags().Int("height", 8, "Chart height")
	excelCmd.AddCommand(excelChartCmd)

	excelAddImageCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelAddImageCmd.Flags().String("cell", "", "Anchor cell for the image (e.g. B2) (required)")
	excelAddImageCmd.Flags().String("image", "", "Path to the image file (required)")
	excelCmd.AddCommand(excelAddImageCmd)

	excelCondFormatCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelCondFormatCmd.Flags().String("range", "", "Cell range (e.g. A1:A10) (required)")
	excelCondFormatCmd.Flags().String("type", "", "Type: cell, top, bottom, average, duplicate, unique, formula (required)")
	excelCondFormatCmd.Flags().String("criteria", "", "Criteria for cell type: >, <, >=, <=, ==, !=")
	excelCondFormatCmd.Flags().String("value", "", "Comparison value")
	excelCondFormatCmd.Flags().String("bg-color", "", "Background color hex")
	excelCondFormatCmd.Flags().String("text-color", "", "Text color hex")
	excelCmd.AddCommand(excelCondFormatCmd)

	excelAutoFilterCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelAutoFilterCmd.Flags().String("range", "", "Header row range (e.g. A1:D1) (required)")
	excelCmd.AddCommand(excelAutoFilterCmd)

	excelMultiChartCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelMultiChartCmd.Flags().String("cell", "", "Anchor cell (e.g. E2) — required")
	excelMultiChartCmd.Flags().String("type", "col", "Chart type: col, bar, line, area, pie, scatter")
	excelMultiChartCmd.Flags().String("spec", "", "JSON spec: {type,cell,series:[{name,catRange,valRange}]} (or @file.json)")
	excelMultiChartCmd.Flags().String("title", "", "Chart title")
	excelMultiChartCmd.Flags().Int("width", 12, "Chart width")
	excelMultiChartCmd.Flags().Int("height", 8, "Chart height")
	excelCmd.AddCommand(excelMultiChartCmd)

	excelDataBarCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelDataBarCmd.Flags().String("range", "", "Cell range (e.g. B2:B100) — required")
	excelDataBarCmd.Flags().String("min-color", "638EC6", "Min bar color hex")
	excelDataBarCmd.Flags().String("max-color", "638EC6", "Max bar color hex")
	excelDataBarCmd.Flags().Bool("bar-only", false, "Hide cell values, show bars only")
	excelCmd.AddCommand(excelDataBarCmd)

	excelIconSetCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelIconSetCmd.Flags().String("range", "", "Cell range (e.g. B2:B100) — required")
	excelIconSetCmd.Flags().String("style", "3Arrows", "Icon style: 3Arrows, 3TrafficLights1, 3TrafficLights2, 3Stars, 4Arrows, 5Arrows, etc.")
	excelIconSetCmd.Flags().Bool("reverse", false, "Reverse icon order")
	excelCmd.AddCommand(excelIconSetCmd)

	excelColorScaleCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelColorScaleCmd.Flags().String("range", "", "Cell range (e.g. B2:B100) — required")
	excelColorScaleCmd.Flags().String("min-color", "F8696B", "Min color hex (default: red)")
	excelColorScaleCmd.Flags().String("max-color", "63BE7B", "Max color hex (default: green)")
	excelColorScaleCmd.Flags().String("mid-color", "", "Mid color hex (omit for 2-color scale)")
	excelColorScaleCmd.Flags().String("min-type", "min", "Min type: min, num, percent, percentile, formula")
	excelColorScaleCmd.Flags().String("max-type", "max", "Max type: max, num, percent, percentile, formula")
	excelColorScaleCmd.Flags().String("min-value", "0", "Min value")
	excelColorScaleCmd.Flags().String("max-value", "0", "Max value")
	excelColorScaleCmd.Flags().String("mid-value", "50", "Mid value (3-color only)")
	excelCmd.AddCommand(excelColorScaleCmd)

	excelDefaultFontCmd.Flags().String("set", "", "Set the default font (omit to just read)")
	excelCmd.AddCommand(excelDefaultFontCmd)

	excelCellStyleCmd.Flags().String("sheet", "", "Sheet name (default: first sheet)")
	excelCmd.AddCommand(excelCellStyleCmd)

	markWrite(excelStyleCmd)
	markWrite(excelSortCmd)
	markWrite(excelFreezeCmd)
	markWrite(excelMergeCmd)
	markWrite(excelUnmergeCmd)
	markWrite(excelSetColWidthCmd)
	markWrite(excelSetRowHeightCmd)
	markWrite(excelChartCmd)
	markWrite(excelAddImageCmd)
	markWrite(excelCondFormatCmd)
	markWrite(excelAutoFilterCmd)
	markWrite(excelMultiChartCmd)
	markWrite(excelDataBarCmd)
	markWrite(excelIconSetCmd)
	markWrite(excelColorScaleCmd)
	markWrite(excelDefaultFontCmd)
}
