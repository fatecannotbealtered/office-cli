package cmd

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/pdf"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var pdfCmd = &cobra.Command{
	Use:   "pdf",
	Short: "Read text, merge, split, trim and watermark PDF files",
}

func init() {
	rootCmd.AddCommand(pdfCmd)

	pdfReadCmd.Flags().Int("page", 0, "Read a single page (1-based); 0 = all")
	pdfReadCmd.Flags().Int("from", 0, "Start page (1-based)")
	pdfReadCmd.Flags().Int("to", 0, "End page (inclusive)")
	pdfReadCmd.Flags().Int("limit", 0, "Max pages to return (0 = all)")
	pdfReadCmd.Flags().Bool("text-only", false, "Output a single concatenated text string instead of per-page array")
	pdfCmd.AddCommand(pdfReadCmd)

	pdfPagesCmd.Flags().Bool("dimensions", false, "Include each page's width/height in points")
	pdfCmd.AddCommand(pdfPagesCmd)

	pdfMergeCmd.Flags().StringSlice("input", nil, "Input PDF paths (repeatable, in order)")
	pdfMergeCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfMergeCmd)

	pdfSplitCmd.Flags().Int("span", 1, "Pages per output file")
	pdfSplitCmd.Flags().String("output-dir", "", "Output directory (required)")
	pdfCmd.AddCommand(pdfSplitCmd)

	pdfTrimCmd.Flags().String("pages", "", "Pages to keep, e.g. '1,3,5-7' (required)")
	pdfTrimCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfTrimCmd)

	pdfWatermarkCmd.Flags().String("text", "", "Watermark text (required)")
	pdfWatermarkCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfWatermarkCmd.Flags().String("style", "font:Helvetica, points:48, opacity:0.3, rotation:45, color:0.7 0.7 0.7", "pdfcpu watermark style descriptor")
	pdfCmd.AddCommand(pdfWatermarkCmd)

	pdfCmd.AddCommand(pdfInfoCmd)

	pdfRotateCmd.Flags().Int("degrees", 90, "Rotation in degrees (90, 180 or 270)")
	pdfRotateCmd.Flags().String("pages", "", "Pages to rotate (e.g. '1,3,5-7'); empty = all pages")
	pdfRotateCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfRotateCmd)

	pdfOptimizeCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfOptimizeCmd)

	pdfEncryptCmd.Flags().String("user-password", "", "User password (required)")
	pdfEncryptCmd.Flags().String("owner-password", "", "Owner password (defaults to user password)")
	pdfEncryptCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfEncryptCmd)

	pdfDecryptCmd.Flags().String("user-password", "", "User password")
	pdfDecryptCmd.Flags().String("owner-password", "", "Owner password")
	pdfDecryptCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfDecryptCmd)

	pdfExtractImagesCmd.Flags().String("output-dir", "", "Output directory (required)")
	pdfExtractImagesCmd.Flags().String("pages", "", "Restrict extraction to these pages (e.g. '1,3,5-7'); empty = all")
	pdfCmd.AddCommand(pdfExtractImagesCmd)

	markWrite(pdfMergeCmd)
	markWrite(pdfSplitCmd)
	markWrite(pdfTrimCmd)
	markWrite(pdfWatermarkCmd)
	markWrite(pdfRotateCmd)
	markWrite(pdfOptimizeCmd)
	markFull(pdfEncryptCmd)
	markFull(pdfDecryptCmd)
	markWrite(pdfExtractImagesCmd)

	// --- New commands ---
	pdfCmd.AddCommand(pdfBookmarksCmd)

	pdfAddBookmarksCmd.Flags().String("spec", "", "JSON array of bookmarks [{title,page,kids:[...]},...] or @file.json (required)")
	pdfAddBookmarksCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfAddBookmarksCmd.Flags().Bool("replace", false, "Replace existing bookmarks")
	pdfCmd.AddCommand(pdfAddBookmarksCmd)

	pdfReorderCmd.Flags().String("order", "", "New page order e.g. '3,1,2,4' (required)")
	pdfReorderCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfReorderCmd)

	pdfInsertBlankCmd.Flags().Int("after", 0, "Insert after this page (1-based) (required)")
	pdfInsertBlankCmd.Flags().Int("count", 1, "Number of blank pages to insert")
	pdfInsertBlankCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfInsertBlankCmd)

	pdfStampImageCmd.Flags().String("image", "", "Path to the image file (required)")
	pdfStampImageCmd.Flags().String("pages", "", "Pages to stamp (e.g. '1,3,5-7'); empty = all")
	pdfStampImageCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfStampImageCmd)

	pdfSetMetaCmd.Flags().String("title", "", "Document title")
	pdfSetMetaCmd.Flags().String("author", "", "Document author")
	pdfSetMetaCmd.Flags().String("subject", "", "Document subject")
	pdfSetMetaCmd.Flags().String("keywords", "", "Document keywords")
	pdfSetMetaCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfSetMetaCmd)

	markWrite(pdfAddBookmarksCmd)
	markWrite(pdfReorderCmd)
	markWrite(pdfInsertBlankCmd)
	markWrite(pdfStampImageCmd)
	markWrite(pdfSetMetaCmd)

	// --- search / create / replace ---
	pdfSearchCmd.Flags().Int("page", 0, "Restrict to a single page (1-based); 0 = all")
	pdfSearchCmd.Flags().Bool("case-sensitive", false, "Case-sensitive match")
	pdfSearchCmd.Flags().Int("context", 30, "Words of surrounding context in snippet")
	pdfSearchCmd.Flags().Int("limit", 100, "Max results to return")
	pdfCmd.AddCommand(pdfSearchCmd)
	// pdfSearchCmd is read-only (no markWrite needed)

	pdfCreateCmd.Flags().String("spec", "", "JSON spec {pages:[{text:...}],title?,author?,paper?} or @file.json (required)")
	pdfCmd.AddCommand(pdfCreateCmd)
	markWrite(pdfCreateCmd)

	pdfReplaceCmd.Flags().String("find", "", "Text to find (single replacement)")
	pdfReplaceCmd.Flags().String("replace", "", "Replacement text")
	pdfReplaceCmd.Flags().String("pairs", "", "JSON array [{find,replace},...] for batch replacement (or @file.json)")
	pdfReplaceCmd.Flags().String("output", "", "Output PDF path (default: <input>.replaced.pdf)")
	pdfCmd.AddCommand(pdfReplaceCmd)
	markWrite(pdfReplaceCmd)

	pdfAddTextCmd.Flags().String("text", "", "Text to overlay (required unless --spec)")
	pdfAddTextCmd.Flags().Float64("x", 72, "X coordinate in points (from left)")
	pdfAddTextCmd.Flags().Float64("y", 100, "Y coordinate in points (from bottom)")
	pdfAddTextCmd.Flags().Float64("font-size", 12, "Font size in points")
	pdfAddTextCmd.Flags().String("color", "", "Text color as 'R G B' (0-1 range, e.g. '1 0 0' for red)")
	pdfAddTextCmd.Flags().String("font", "Helvetica", "Font family: Helvetica, Times, Courier")
	pdfAddTextCmd.Flags().Bool("bold", false, "Use bold font variant")
	pdfAddTextCmd.Flags().Bool("italic", false, "Use italic font variant")
	pdfAddTextCmd.Flags().Bool("underline", false, "Add underline decoration")
	pdfAddTextCmd.Flags().String("pages", "", "Pages to overlay (e.g. '1,3,5-7'); empty = all")
	pdfAddTextCmd.Flags().String("spec", "", "JSON array [{text,x,y,fontSize?,color?,font?,bold?,italic?,underline?,pages?},...] or @file.json")
	pdfAddTextCmd.Flags().String("output", "", "Output PDF path (required)")
	pdfCmd.AddCommand(pdfAddTextCmd)
	markWrite(pdfAddTextCmd)
}

var pdfReadCmd = &cobra.Command{
	Use:   "read <FILE>",
	Short: "Extract text from a PDF (per-page or single string)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		page, _ := cmd.Flags().GetInt("page")
		from, _ := cmd.Flags().GetInt("from")
		to, _ := cmd.Flags().GetInt("to")
		limit, _ := cmd.Flags().GetInt("limit")
		textOnly, _ := cmd.Flags().GetBool("text-only")

		if page > 0 {
			from = page
			to = page
		}

		if textOnly {
			text, err := pdf.AllText(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "text": text})
				return nil
			}
			fmt.Println(text)
			return nil
		}

		pages, err := pdf.Read(path, pdf.ReadOptions{From: from, To: to, Limit: limit})
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			flat := make([]output.FlatPage, 0, len(pages))
			for _, p := range pages {
				flat = append(flat, output.FlatPage{Page: p.Page, Text: p.Text, WordCount: p.WordCount})
			}
			output.PrintJSON(map[string]any{"file": path, "pages": flat})
			return nil
		}

		for _, p := range pages {
			output.Bold(fmt.Sprintf("--- page %d (%d words) ---", p.Page, p.WordCount))
			fmt.Println(p.Text)
			fmt.Println()
		}
		return nil
	},
}

var pdfPagesCmd = &cobra.Command{
	Use:   "pages <FILE>",
	Short: "Show page count (and optionally page dimensions)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		count, err := pdf.PageCount(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		dimensionsRequested, _ := cmd.Flags().GetBool("dimensions")

		if dimensionsRequested {
			dims, err := pdf.PageDimensions(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "pages": count, "dimensions": dims})
				return nil
			}
			rows := make([][]string, 0, len(dims))
			for _, d := range dims {
				rows = append(rows, []string{
					fmt.Sprintf("%d", d.Page),
					fmt.Sprintf("%.1f", d.Width),
					fmt.Sprintf("%.1f", d.Height),
				})
			}
			output.Table([]string{"page", "width(pt)", "height(pt)"}, rows)
			return nil
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "pages": count})
			return nil
		}
		fmt.Printf("%d\n", count)
		return nil
	},
}

var pdfMergeCmd = &cobra.Command{
	Use:   "merge [FILE...]",
	Short: "Concatenate multiple PDFs into one",
	Long: `Concatenate multiple PDFs into one. Accepts inputs as positional
arguments or via repeated --input flags. The order on the command line is
preserved in the output document.`,
	Args: cobra.ArbitraryArgs,
	RunE: func(cmd *cobra.Command, args []string) error {
		inputs, _ := cmd.Flags().GetStringSlice("input")
		inputs = append(inputs, args...)
		out, _ := cmd.Flags().GetString("output")
		if len(inputs) < 2 {
			return emitError("at least 2 input files are required (use positional args or repeated --input)", output.ErrValidation, "", ExitBadArgs)
		}
		for i, p := range inputs {
			normalized, err := resolveInput(p, "pdf")
			if err != nil {
				return err
			}
			inputs[i] = normalized
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err := resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("merge PDFs", map[string]any{"inputs": inputs, "output": out}) {
			return nil
		}

		if err := pdf.Merge(inputs, out); err != nil {
			return emitError(err.Error(), output.ErrEngine, out, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "output": out, "inputs": inputs})
			return nil
		}
		output.Success(fmt.Sprintf("merged %d files -> %s", len(inputs), out))
		return nil
	},
}

var pdfSplitCmd = &cobra.Command{
	Use:   "split <FILE>",
	Short: "Split PDF into chunks of N pages each",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		span, _ := cmd.Flags().GetInt("span")
		outDir, _ := cmd.Flags().GetString("output-dir")
		if outDir == "" {
			return emitError("--output-dir is required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("split PDF", map[string]any{"file": path, "span": span, "outputDir": outDir}) {
			return nil
		}

		if err := pdf.SplitEvery(path, outDir, span); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		var produced []string
		_ = filepath.Walk(outDir, func(p string, info os.FileInfo, err error) error {
			if err == nil && !info.IsDir() && strings.EqualFold(filepath.Ext(p), ".pdf") {
				produced = append(produced, p)
			}
			return nil
		})

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "outputDir": outDir, "files": produced})
			return nil
		}
		output.Success(fmt.Sprintf("split %s into %d file(s) under %s", path, len(produced), outDir))
		return nil
	},
}

var pdfTrimCmd = &cobra.Command{
	Use:   "trim <FILE>",
	Short: "Keep only the listed pages (e.g. --pages '1,3,5-7') into a new file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		pages, _ := cmd.Flags().GetString("pages")
		out, _ := cmd.Flags().GetString("output")
		if pages == "" {
			return emitError("--pages is required (e.g. 1,3,5-7)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		if _, err := pdf.ParsePageList(pages); err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("trim PDF", map[string]any{"file": path, "pages": pages, "output": out}) {
			return nil
		}

		if err := pdf.Trim(path, out, pages); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "pages": pages, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("trimmed pages %s of %s -> %s", pages, path, out))
		return nil
	},
}

var pdfWatermarkCmd = &cobra.Command{
	Use:   "watermark <FILE>",
	Short: "Add a text watermark to every page",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		text, _ := cmd.Flags().GetString("text")
		out, _ := cmd.Flags().GetString("output")
		style, _ := cmd.Flags().GetString("style")
		if text == "" {
			return emitError("--text is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}

		if dryRunOutput("add watermark", map[string]any{"file": path, "text": text, "output": out}) {
			return nil
		}

		if err := pdf.WatermarkText(path, out, text, style); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out, "text": text})
			return nil
		}
		output.Success(fmt.Sprintf("watermarked %s -> %s", path, out))
		return nil
	},
}

var pdfInfoCmd = &cobra.Command{
	Use:   "info <FILE>",
	Short: "Show full PDF metadata (title, author, encryption, signatures, dimensions, ...)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		info, err := pdf.FullInfo(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(info)
			return nil
		}
		rows := [][]string{
			{"path", fmt.Sprintf("%v", info["path"])},
			{"sizeBytes", fmt.Sprintf("%v", info["sizeBytes"])},
			{"version", fmt.Sprintf("%v", info["version"])},
			{"pageCount", fmt.Sprintf("%v", info["pageCount"])},
			{"title", fmt.Sprintf("%v", info["title"])},
			{"author", fmt.Sprintf("%v", info["author"])},
			{"subject", fmt.Sprintf("%v", info["subject"])},
			{"creator", fmt.Sprintf("%v", info["creator"])},
			{"producer", fmt.Sprintf("%v", info["producer"])},
			{"creationDate", fmt.Sprintf("%v", info["creationDate"])},
			{"modificationDate", fmt.Sprintf("%v", info["modificationDate"])},
			{"encrypted", fmt.Sprintf("%v", info["encrypted"])},
			{"watermarked", fmt.Sprintf("%v", info["watermarked"])},
			{"linearized", fmt.Sprintf("%v", info["linearized"])},
			{"tagged", fmt.Sprintf("%v", info["tagged"])},
			{"form", fmt.Sprintf("%v", info["form"])},
			{"signatures", fmt.Sprintf("%v", info["signatures"])},
			{"bookmarks", fmt.Sprintf("%v", info["bookmarks"])},
		}
		output.Table([]string{"field", "value"}, rows)
		return nil
	},
}

var pdfRotateCmd = &cobra.Command{
	Use:   "rotate <FILE>",
	Short: "Rotate selected pages by 90 / 180 / 270 degrees into a new file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		degrees, _ := cmd.Flags().GetInt("degrees")
		pages, _ := cmd.Flags().GetString("pages")
		out, _ := cmd.Flags().GetString("output")
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}
		if pages != "" {
			if _, err := pdf.ParsePageList(pages); err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}

		if dryRunOutput("rotate PDF", map[string]any{"file": path, "degrees": degrees, "pages": pages, "output": out}) {
			return nil
		}

		if err := pdf.Rotate(path, out, degrees, pages); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "degrees": degrees, "pages": pages, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("rotated %s by %d° -> %s", path, degrees, out))
		return nil
	},
}

var pdfOptimizeCmd = &cobra.Command{
	Use:   "optimize <FILE>",
	Short: "Reduce file size by compacting cross-references and reusing objects",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		out, _ := cmd.Flags().GetString("output")
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("optimize PDF", map[string]any{"file": path, "output": out}) {
			return nil
		}

		if err := pdf.Optimize(path, out); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		inSize, _ := os.Stat(path)
		outSize, _ := os.Stat(out)
		if jsonMode {
			payload := map[string]any{"status": "ok", "file": path, "output": out}
			if inSize != nil && outSize != nil {
				payload["sizeBefore"] = inSize.Size()
				payload["sizeAfter"] = outSize.Size()
				if inSize.Size() > 0 {
					payload["savedPercent"] = 100 * float64(inSize.Size()-outSize.Size()) / float64(inSize.Size())
				}
			}
			output.PrintJSON(payload)
			return nil
		}
		if inSize != nil && outSize != nil {
			output.Success(fmt.Sprintf("optimized %s (%d -> %d bytes) -> %s", path, inSize.Size(), outSize.Size(), out))
		} else {
			output.Success(fmt.Sprintf("optimized %s -> %s", path, out))
		}
		return nil
	},
}

var pdfEncryptCmd = &cobra.Command{
	Use:   "encrypt <FILE>",
	Short: "Encrypt a PDF with a user (and optional owner) password",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		userPW, _ := cmd.Flags().GetString("user-password")
		ownerPW, _ := cmd.Flags().GetString("owner-password")
		out, _ := cmd.Flags().GetString("output")
		if userPW == "" {
			return emitError("--user-password is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("encrypt PDF", map[string]any{"file": path, "output": out}) {
			return nil
		}

		if err := pdf.Encrypt(path, out, userPW, ownerPW); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("encrypted %s -> %s", path, out))
		return nil
	},
}

var pdfDecryptCmd = &cobra.Command{
	Use:   "decrypt <FILE>",
	Short: "Strip encryption from a PDF when given the right password",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		userPW, _ := cmd.Flags().GetString("user-password")
		ownerPW, _ := cmd.Flags().GetString("owner-password")
		out, _ := cmd.Flags().GetString("output")
		if userPW == "" && ownerPW == "" {
			return emitError("provide --user-password or --owner-password", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("decrypt PDF", map[string]any{"file": path, "output": out}) {
			return nil
		}

		if err := pdf.Decrypt(path, out, userPW, ownerPW); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("decrypted %s -> %s", path, out))
		return nil
	},
}

var pdfExtractImagesCmd = &cobra.Command{
	Use:   "extract-images <FILE>",
	Short: "Extract every image embedded in the PDF into --output-dir",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		outDir, _ := cmd.Flags().GetString("output-dir")
		pages, _ := cmd.Flags().GetString("pages")
		if outDir == "" {
			return emitError("--output-dir is required", output.ErrValidation, "", ExitBadArgs)
		}
		outDir, err = resolveOutput(outDir)
		if err != nil {
			return err
		}
		if pages != "" {
			if _, err := pdf.ParsePageList(pages); err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}

		if dryRunOutput("extract images", map[string]any{"file": path, "outputDir": outDir, "pages": pages}) {
			return nil
		}

		if err := pdf.ExtractImages(path, outDir, pages); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		images, _ := pdf.ListExtractedImages(outDir)
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "outputDir": outDir, "images": images})
			return nil
		}
		if len(images) == 0 {
			output.Gray("(no images extracted)")
			return nil
		}
		rows := make([][]string, 0, len(images))
		for _, img := range images {
			rows = append(rows, []string{img.Name, fmt.Sprintf("%d", img.Bytes), img.Path})
		}
		output.Table([]string{"name", "bytes", "path"}, rows)
		return nil
	},
}

// ---------------------------------------------------------------------------
// New PDF commands: bookmarks, reorder, insert-blank, stamp-image, set-meta
// ---------------------------------------------------------------------------

var pdfBookmarksCmd = &cobra.Command{
	Use:   "bookmarks <FILE>",
	Short: "Read the PDF outline/bookmarks tree",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		bms, err := pdf.Bookmarks(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "bookmarks": bms})
			return nil
		}
		if len(bms) == 0 {
			output.Gray("(no bookmarks)")
			return nil
		}
		printBookmarkTree(bms, 0)
		return nil
	},
}

func printBookmarkTree(bms []pdf.Bookmark, depth int) {
	for _, b := range bms {
		indent := strings.Repeat("  ", depth)
		fmt.Printf("%s%d. %s (page %d)\n", indent, depth+1, b.Title, b.Page)
		if len(b.Children) > 0 {
			printBookmarkTree(b.Children, depth+1)
		}
	}
}

var pdfAddBookmarksCmd = &cobra.Command{
	Use:   "add-bookmarks <FILE>",
	Short: "Add bookmarks (outline entries) to a PDF",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		specRaw, _ := cmd.Flags().GetString("spec")
		out, _ := cmd.Flags().GetString("output")
		replace, _ := cmd.Flags().GetBool("replace")
		if specRaw == "" {
			return emitError("--spec is required (JSON array, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var specs []pdf.BookmarkSpec
		if err := json.Unmarshal(data, &specs); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("add bookmarks", map[string]any{"file": path, "count": len(specs), "output": out}) {
			return nil
		}

		if err := pdf.AddBookmarks(path, out, specs, replace); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out, "bookmarks": len(specs)})
			return nil
		}
		output.Success(fmt.Sprintf("added %d bookmark(s) to %s -> %s", len(specs), path, out))
		return nil
	},
}

var pdfReorderCmd = &cobra.Command{
	Use:   "reorder <FILE>",
	Short: "Reorder pages in a PDF",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		order, _ := cmd.Flags().GetString("order")
		out, _ := cmd.Flags().GetString("output")
		if order == "" {
			return emitError("--order is required (e.g. '3,1,2,4')", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("reorder pages", map[string]any{"file": path, "order": order, "output": out}) {
			return nil
		}

		if err := pdf.Reorder(path, out, order); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "order": order, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("reordered pages of %s -> %s", path, out))
		return nil
	},
}

var pdfInsertBlankCmd = &cobra.Command{
	Use:   "insert-blank <FILE>",
	Short: "Insert blank pages into a PDF",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		after, _ := cmd.Flags().GetInt("after")
		count, _ := cmd.Flags().GetInt("count")
		out, _ := cmd.Flags().GetString("output")
		if after < 1 {
			return emitError("--after must be >= 1", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("insert blank pages", map[string]any{"file": path, "after": after, "count": count, "output": out}) {
			return nil
		}

		if err := pdf.InsertBlank(path, out, after, count); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "after": after, "count": count, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("inserted %d blank page(s) after page %d in %s -> %s", count, after, path, out))
		return nil
	},
}

var pdfStampImageCmd = &cobra.Command{
	Use:   "stamp-image <FILE>",
	Short: "Stamp an image on PDF pages",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		imagePath, _ := cmd.Flags().GetString("image")
		pages, _ := cmd.Flags().GetString("pages")
		out, _ := cmd.Flags().GetString("output")
		if imagePath == "" {
			return emitError("--image is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("stamp image", map[string]any{"file": path, "image": imagePath, "output": out}) {
			return nil
		}

		if err := pdf.StampImage(path, out, imagePath, pages); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "image": imagePath, "output": out})
			return nil
		}
		output.Success(fmt.Sprintf("stamped image on %s -> %s", path, out))
		return nil
	},
}

var pdfSetMetaCmd = &cobra.Command{
	Use:   "set-meta <FILE>",
	Short: "Set PDF metadata (title, author, subject, keywords)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		title, _ := cmd.Flags().GetString("title")
		author, _ := cmd.Flags().GetString("author")
		subject, _ := cmd.Flags().GetString("subject")
		keywords, _ := cmd.Flags().GetString("keywords")
		out, _ := cmd.Flags().GetString("output")
		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		meta := pdf.MetaUpdate{Title: title, Author: author, Subject: subject, Keywords: keywords}

		if dryRunOutput("set metadata", map[string]any{"file": path, "output": out}) {
			return nil
		}

		if err := pdf.SetMeta(path, out, meta); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out, "title": title, "author": author})
			return nil
		}
		output.Success(fmt.Sprintf("set metadata on %s -> %s", path, out))
		return nil
	},
}

// ---------------------------------------------------------------------------
// New PDF commands: search, create, replace
// ---------------------------------------------------------------------------

var pdfSearchCmd = &cobra.Command{
	Use:   "search <FILE> <KEYWORD>",
	Short: "Find text occurrences in a PDF (case-insensitive by default)",
	Args:  cobra.ExactArgs(2),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		keyword := args[1]
		page, _ := cmd.Flags().GetInt("page")
		caseSensitive, _ := cmd.Flags().GetBool("case-sensitive")
		context, _ := cmd.Flags().GetInt("context")
		limit, _ := cmd.Flags().GetInt("limit")

		results, err := pdf.Search(path, keyword, pdf.SearchOptions{
			Page:          page,
			CaseSensitive: caseSensitive,
			Context:       context,
			Limit:         limit,
		})
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "keyword": keyword, "count": len(results), "matches": results})
			return nil
		}
		if len(results) == 0 {
			output.Gray("(no matches)")
			return nil
		}
		rows := make([][]string, 0, len(results))
		for _, r := range results {
			rows = append(rows, []string{
				fmt.Sprintf("%d", r.Page),
				fmt.Sprintf("%d", r.Line),
				r.Snippet,
			})
		}
		output.Table([]string{"page", "line", "snippet"}, rows)
		fmt.Printf("\n%d match(es) found\n", len(results))
		return nil
	},
}

var pdfCreateCmd = &cobra.Command{
	Use:   "create <OUTPUT>",
	Short: "Create a new PDF from a JSON spec",
	Long: "Create a new PDF with one or more pages of text content.\n\n" +
		"The --spec flag accepts inline JSON or @file.json:\n\n" +
		`  --spec '{"pages":[{"text":"Hello World"},{"text":"Page 2 content"}]}'`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		out := args[0]
		specRaw, _ := cmd.Flags().GetString("spec")
		if specRaw == "" {
			return emitError("--spec is required (JSON object, or @file.json)", output.ErrValidation, "", ExitBadArgs)
		}
		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		var spec pdf.CreateSpec
		if err := json.Unmarshal(data, &spec); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
		}
		if len(spec.Pages) == 0 {
			return emitError("--spec must contain at least one page", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("create PDF", map[string]any{"output": out, "pages": len(spec.Pages)}) {
			return nil
		}

		if err := pdf.CreatePDF(out, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, out, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "output": out, "pages": len(spec.Pages)})
			return nil
		}
		output.Success(fmt.Sprintf("created %s (%d page(s))", out, len(spec.Pages)))
		return nil
	},
}

var pdfReplaceCmd = &cobra.Command{
	Use:   "replace <FILE>",
	Short: "Find and replace text in a PDF",
	Long: `Find and replace text literals in the PDF content stream.

Single replacement:
  pdf replace input.pdf --find "Old" --replace "New" --output out.pdf

Batch replacement:
  pdf replace input.pdf --pairs '[{"find":"X","replace":"Y"},{"find":"A","replace":"B"}]' --output out.pdf

Limitation: only replaces text stored as literal strings in Tj operators.
Text split across operators or using hex strings is not affected.`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		find, _ := cmd.Flags().GetString("find")
		repl, _ := cmd.Flags().GetString("replace")
		pairsRaw, _ := cmd.Flags().GetString("pairs")
		out, _ := cmd.Flags().GetString("output")

		var pairs []pdf.Replacement
		if pairsRaw != "" {
			data, err := readSpecArg(pairsRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &pairs); err != nil {
				return emitError("invalid --pairs JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		} else if find != "" {
			pairs = []pdf.Replacement{{Find: find, Replace: repl}}
		} else {
			return emitError("provide --find/--replace or --pairs", output.ErrValidation, "", ExitBadArgs)
		}

		if out == "" {
			ext := filepath.Ext(path)
			base := strings.TrimSuffix(path, ext)
			out = base + ".replaced" + ext
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("replace text in PDF", map[string]any{"file": path, "pairs": len(pairs), "output": out}) {
			return nil
		}

		hits, err := pdf.ReplaceText(path, out, pairs)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			payload := map[string]any{"status": "ok", "file": path, "output": out, "hits": hits}
			if hits == 0 {
				payload["warning"] = "no replacements made — text may be split across operators or stored as hex strings; only parenthesized Tj literals are matched"
			}
			output.PrintJSON(payload)
			return nil
		}
		if hits == 0 {
			output.Warn("no replacements made — text may be split across Tj operators or stored as hex strings")
		}
		output.Success(fmt.Sprintf("replaced text in %s (%d hit(s)) -> %s", path, hits, out))
		return nil
	},
}

var pdfAddTextCmd = &cobra.Command{
	Use:   "add-text <FILE>",
	Short: "Overlay text at specific coordinates on PDF pages",
	Long: `Overlay text at precise (x, y) coordinates on PDF pages.

Single overlay via flags:
  pdf add-text input.pdf --text "DRAFT" --x 200 --y 400 --font-size 48 --color "1 0 0" --output out.pdf

Batch overlays via --spec:
  pdf add-text input.pdf --spec '[{"text":"Page 1","x":72,"y":700},{"text":"Confidential","x":200,"y":400,"fontSize":24,"color":"1 0 0","pages":"1"}]' --output out.pdf`,
	Args: cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pdf")
		if err != nil {
			return err
		}
		specRaw, _ := cmd.Flags().GetString("spec")
		text, _ := cmd.Flags().GetString("text")
		x, _ := cmd.Flags().GetFloat64("x")
		y, _ := cmd.Flags().GetFloat64("y")
		fontSize, _ := cmd.Flags().GetFloat64("font-size")
		color, _ := cmd.Flags().GetString("color")
		font, _ := cmd.Flags().GetString("font")
		bold, _ := cmd.Flags().GetBool("bold")
		italic, _ := cmd.Flags().GetBool("italic")
		underline, _ := cmd.Flags().GetBool("underline")
		pages, _ := cmd.Flags().GetString("pages")
		out, _ := cmd.Flags().GetString("output")

		var overlays []pdf.TextOverlay
		if specRaw != "" {
			data, err := readSpecArg(specRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &overlays); err != nil {
				return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		} else if text != "" {
			overlays = []pdf.TextOverlay{{
				Text:      text,
				X:         x,
				Y:         y,
				FontSize:  fontSize,
				Bold:      bold,
				Italic:    italic,
				Underline: underline,
				Color:     color,
				Font:      font,
				Pages:     pages,
			}}
		} else {
			return emitError("provide --text or --spec", output.ErrValidation, "", ExitBadArgs)
		}

		if out == "" {
			return emitError("--output is required", output.ErrValidation, "", ExitBadArgs)
		}
		out, err = resolveOutput(out)
		if err != nil {
			return err
		}

		if dryRunOutput("add text overlay", map[string]any{"file": path, "overlays": len(overlays), "output": out}) {
			return nil
		}

		if err := pdf.AddText(path, out, overlays); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "output": out, "overlays": len(overlays)})
			return nil
		}
		output.Success(fmt.Sprintf("added %d text overlay(s) to %s -> %s", len(overlays), path, out))
		return nil
	},
}
