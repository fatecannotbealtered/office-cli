package cmd

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/word"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// Word write commands: create, add-paragraph, add-heading, add-table,
// add-image, add-page-break, replace, merge
// ---------------------------------------------------------------------------

var wordCreateCmd = &cobra.Command{
	Use:   "create <FILE>",
	Short: "Create a new .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		title, _ := cmd.Flags().GetString("title")
		author, _ := cmd.Flags().GetString("author")

		if dryRunOutput("create document", map[string]any{"file": path, "title": title, "author": author}) {
			return nil
		}

		if err := word.Create(path, title, author); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path})
			return nil
		}
		output.Success(fmt.Sprintf("created %s", path))
		return nil
	},
}

var wordAddParagraphCmd = &cobra.Command{
	Use:   "add-paragraph <FILE>",
	Short: "Append a paragraph to a .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		text, _ := cmd.Flags().GetString("text")
		style, _ := cmd.Flags().GetString("style")
		out, _ := cmd.Flags().GetString("output")
		if text == "" {
			return emitError("--text is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add paragraph", map[string]any{"file": path, "style": style}) {
			return nil
		}

		if err := word.AddParagraph(path, out, text, style); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "style": style})
			return nil
		}
		output.Success(fmt.Sprintf("added paragraph to %s", out))
		return nil
	},
}

var wordAddHeadingCmd = &cobra.Command{
	Use:   "add-heading <FILE>",
	Short: "Append a heading to a .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		text, _ := cmd.Flags().GetString("text")
		level, _ := cmd.Flags().GetInt("level")
		out, _ := cmd.Flags().GetString("output")
		if text == "" {
			return emitError("--text is required", output.ErrValidation, "", ExitBadArgs)
		}
		if level < 1 || level > 6 {
			return emitError("--level must be 1-6", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add heading", map[string]any{"file": path, "level": level}) {
			return nil
		}

		if err := word.AddHeading(path, out, text, level); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "level": level})
			return nil
		}
		output.Success(fmt.Sprintf("added Heading%d to %s", level, out))
		return nil
	},
}

var wordAddTableCmd = &cobra.Command{
	Use:   "add-table <FILE>",
	Short: "Append a table to a .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		rowsRaw, _ := cmd.Flags().GetString("rows")
		out, _ := cmd.Flags().GetString("output")
		if rowsRaw == "" {
			return emitError("--rows is required (JSON array of arrays, or @file.json)", output.ErrValidation, "", ExitBadArgs)
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

		if dryRunOutput("add table", map[string]any{"file": path, "rows": len(rows)}) {
			return nil
		}

		if err := word.AddTable(path, out, rows); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "rows": len(rows)})
			return nil
		}
		output.Success(fmt.Sprintf("added %d-row table to %s", len(rows), out))
		return nil
	},
}

var wordAddImageCmd = &cobra.Command{
	Use:   "add-image <FILE>",
	Short: "Insert an inline image into a .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		imagePath, _ := cmd.Flags().GetString("image")
		width, _ := cmd.Flags().GetFloat64("width")
		height, _ := cmd.Flags().GetFloat64("height")
		out, _ := cmd.Flags().GetString("output")
		if imagePath == "" {
			return emitError("--image is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add image", map[string]any{"file": path, "image": imagePath}) {
			return nil
		}

		if err := word.AddImage(path, out, imagePath, width, height); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "image": imagePath})
			return nil
		}
		output.Success(fmt.Sprintf("added image to %s", out))
		return nil
	},
}

var wordAddPageBreakCmd = &cobra.Command{
	Use:   "add-page-break <FILE>",
	Short: "Insert a page break into a .docx document",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		out, _ := cmd.Flags().GetString("output")
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add page break", map[string]any{"file": path}) {
			return nil
		}

		if err := word.AddPageBreak(path, out); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out})
			return nil
		}
		output.Success(fmt.Sprintf("added page break to %s", out))
		return nil
	},
}

var wordReplaceCmd = &cobra.Command{
	Use:   "replace <FILE>",
	Short: "Find/replace text in a .docx (creates a new file, source is untouched)",
	Long: `Performs find-and-replace within each text run of word/document.xml.

Use --find/--replace for a single pair, or --pairs '[{"find":"X","replace":"Y"}, ...]'
(or @file.json) for a batch.

Limitation: replacements that would span multiple runs (e.g. a word split by an
italic boundary) are not detected. The skill documentation explains how AI Agents
should normalize text before relying on replacements.`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		find, _ := cmd.Flags().GetString("find")
		replace, _ := cmd.Flags().GetString("replace")
		pairsRaw, _ := cmd.Flags().GetString("pairs")
		out, _ := cmd.Flags().GetString("output")

		var pairs []word.Replacement
		if pairsRaw != "" {
			data, err := readSpecArg(pairsRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &pairs); err != nil {
				return emitError("invalid --pairs JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}
		if find != "" {
			pairs = append(pairs, word.Replacement{Find: find, Replace: replace})
		}
		if len(pairs) == 0 {
			return emitError("provide --find/--replace or --pairs", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = strings.TrimSuffix(path, ".docx") + ".replaced.docx"
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("replace text", map[string]any{"file": path, "pairs": len(pairs), "output": out}) {
			return nil
		}

		hits, err := word.ReplaceText(path, out, pairs)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			payload := map[string]any{"status": "ok", "file": path, "output": out, "hits": hits}
			if hits == 0 {
				payload["warning"] = "no replacements made — text may be split across multiple runs; only text within a single <w:t> element is matched"
			}
			output.PrintJSON(payload)
			return nil
		}
		if hits == 0 {
			output.Warn("no replacements made — text may be split across multiple runs")
		}
		output.Success(fmt.Sprintf("replaced %d occurrence(s) -> %s", hits, out))
		return nil
	},
}

var wordMergeCmd = &cobra.Command{
	Use:   "merge FILE1 FILE2 [FILE3...]",
	Short: "Merge multiple .docx files into one (--output is required)",
	Args:  cobra.MinimumNArgs(2),
	RunE: func(cmd *cobra.Command, args []string) error {
		out, _ := cmd.Flags().GetString("output")
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		outPath, err := resolveOutput(out)
		if err != nil {
			return err
		}

		// Validate all input files
		var paths []string
		for _, a := range args {
			p, err := resolveInput(a, "docx")
			if err != nil {
				return err
			}
			paths = append(paths, p)
		}

		if dryRunOutput("merge documents", map[string]any{"files": len(paths), "output": outPath}) {
			return nil
		}

		if err := word.Merge(paths, outPath); err != nil {
			return emitError(err.Error(), output.ErrEngine, "", ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "files": len(paths), "output": outPath})
			return nil
		}
		output.Success(fmt.Sprintf("merged %d files -> %s", len(paths), outPath))
		return nil
	},
}

func init() {
	wordCreateCmd.Flags().String("title", "", "Document title")
	wordCreateCmd.Flags().String("author", "", "Document author")
	wordCmd.AddCommand(wordCreateCmd)

	wordAddParagraphCmd.Flags().String("text", "", "Paragraph text (required)")
	wordAddParagraphCmd.Flags().String("style", "", "Style: Normal, Heading1-Heading6, Title")
	wordAddParagraphCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddParagraphCmd)

	wordAddHeadingCmd.Flags().String("text", "", "Heading text (required)")
	wordAddHeadingCmd.Flags().Int("level", 1, "Heading level 1-6")
	wordAddHeadingCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddHeadingCmd)

	wordAddTableCmd.Flags().String("rows", "", "JSON array of arrays (required, or @file.json)")
	wordAddTableCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddTableCmd)

	wordAddImageCmd.Flags().String("image", "", "Path to the image file (required)")
	wordAddImageCmd.Flags().Float64("width", 200, "Image width in points")
	wordAddImageCmd.Flags().Float64("height", 200, "Image height in points")
	wordAddImageCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddImageCmd)

	wordAddPageBreakCmd.Flags().String("output", "", "Output .docx path (defaults to input)")
	wordCmd.AddCommand(wordAddPageBreakCmd)

	wordReplaceCmd.Flags().String("find", "", "Substring to find (use with --replace, repeatable via --pairs)")
	wordReplaceCmd.Flags().String("replace", "", "Replacement string for --find")
	wordReplaceCmd.Flags().String("pairs", "", "JSON array of {find,replace} (or @file.json) for batch replacement")
	wordReplaceCmd.Flags().String("output", "", "Output .docx path (defaults to <input>.replaced.docx)")
	wordCmd.AddCommand(wordReplaceCmd)

	wordMergeCmd.Flags().String("output", "", "Output .docx path (required)")
	wordCmd.AddCommand(wordMergeCmd)

	markWrite(wordReplaceCmd)
	markWrite(wordCreateCmd)
	markWrite(wordAddParagraphCmd)
	markWrite(wordAddHeadingCmd)
	markWrite(wordAddTableCmd)
	markWrite(wordAddImageCmd)
	markWrite(wordAddPageBreakCmd)
	markWrite(wordMergeCmd)
}
