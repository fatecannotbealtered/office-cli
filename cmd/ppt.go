package cmd

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/ppt"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var pptCmd = &cobra.Command{
	Use:   "ppt",
	Short: "Read slide outlines, replace text and inspect metadata of .pptx files",
}

func init() {
	rootCmd.AddCommand(pptCmd)

	pptReadCmd.Flags().Int("slide", 0, "Read a single slide (1-based); 0 = all")
	pptReadCmd.Flags().String("format", "outline", "Output format: outline | markdown | text")
	pptReadCmd.Flags().String("keyword", "", "Only return slides containing this keyword (case-insensitive)")
	pptReadCmd.Flags().Bool("with-notes", false, "Include speaker notes in the output")
	pptCmd.AddCommand(pptReadCmd)

	pptReplaceCmd.Flags().String("find", "", "Substring to find (use with --replace, repeatable via --pairs)")
	pptReplaceCmd.Flags().String("replace", "", "Replacement string for --find")
	pptReplaceCmd.Flags().String("pairs", "", "JSON array of {find,replace} (or @file.json) for batch replacement")
	pptReplaceCmd.Flags().String("output", "", "Output .pptx path (defaults to <input>.replaced.pptx)")
	pptCmd.AddCommand(pptReplaceCmd)

	pptCmd.AddCommand(pptMetaCmd)
	pptCmd.AddCommand(pptCountCmd)
	pptCmd.AddCommand(pptOutlineCmd)

	pptImagesCmd.Flags().String("output-dir", "", "Directory to write extracted images (required)")
	pptCmd.AddCommand(pptImagesCmd)

	// --- New write commands ---
	pptCreateCmd.Flags().String("title", "", "Presentation title")
	pptCreateCmd.Flags().String("author", "", "Presentation author")
	pptCmd.AddCommand(pptCreateCmd)

	pptAddSlideCmd.Flags().String("title", "", "Slide title")
	pptAddSlideCmd.Flags().String("bullets", "", "JSON array of bullet strings (or @file.json)")
	pptAddSlideCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptAddSlideCmd)

	pptSetContentCmd.Flags().Int("slide", 0, "Slide number (1-based, required)")
	pptSetContentCmd.Flags().String("title", "", "New slide title")
	pptSetContentCmd.Flags().String("bullets", "", "JSON array of bullet strings (or @file.json)")
	pptSetContentCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptSetContentCmd)

	pptSetNotesCmd.Flags().Int("slide", 0, "Slide number (1-based, required)")
	pptSetNotesCmd.Flags().String("notes", "", "Speaker notes text (required)")
	pptSetNotesCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptSetNotesCmd)

	pptDeleteSlideCmd.Flags().Int("slide", 0, "Slide number to delete (1-based, required)")
	pptDeleteSlideCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptDeleteSlideCmd)

	pptReorderCmd.Flags().String("order", "", "Comma-separated new order, e.g. '3,1,2' (required)")
	pptReorderCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptReorderCmd)

	pptAddImageCmd.Flags().Int("slide", 0, "Slide number (1-based, required)")
	pptAddImageCmd.Flags().String("image", "", "Path to the image file (required)")
	pptAddImageCmd.Flags().Int("width", 0, "Image width in EMU (0 = full slide width)")
	pptAddImageCmd.Flags().Int("height", 0, "Image height in EMU (0 = full slide height)")
	pptAddImageCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptAddImageCmd)

	pptBuildCmd.Flags().String("spec", "", "JSON spec (string or @file.json) with title, author, and slides array (required)")
	pptBuildCmd.Flags().String("template", "", "Path to a .pptx template file (preserves slide master, theme, and layouts)")
	pptCmd.AddCommand(pptBuildCmd)

	// --- Phase 3: layout, set-style, add-shape ---
	pptLayoutCmd.Flags().Int("slide", 0, "Slide number (1-based); 0 = all slides")
	pptCmd.AddCommand(pptLayoutCmd)

	pptSetStyleCmd.Flags().Int("slide", 0, "Slide number (1-based, required)")
	pptSetStyleCmd.Flags().Int("shape", 0, "Shape index (0-based, from 'ppt layout', required)")
	pptSetStyleCmd.Flags().Int("font-size", 0, "Font size in hundredths of a point (e.g. 2400 = 24pt)")
	pptSetStyleCmd.Flags().Bool("bold", false, "Set bold")
	pptSetStyleCmd.Flags().Bool("italic", false, "Set italic")
	pptSetStyleCmd.Flags().Bool("underline", false, "Set underline")
	pptSetStyleCmd.Flags().String("color", "", "Text color as hex RGB (e.g. FF0000)")
	pptSetStyleCmd.Flags().String("align", "", "Paragraph alignment: left, right, center, justify")
	pptSetStyleCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptSetStyleCmd)

	pptAddShapeCmd.Flags().Int("slide", 0, "Slide number (1-based, required)")
	pptAddShapeCmd.Flags().String("type", "", "Shape type: text-box, rect, ellipse, line, arrow (required)")
	pptAddShapeCmd.Flags().Int("x", 0, "X position in EMU")
	pptAddShapeCmd.Flags().Int("y", 0, "Y position in EMU")
	pptAddShapeCmd.Flags().Int("width", 0, "Width in EMU")
	pptAddShapeCmd.Flags().Int("height", 0, "Height in EMU")
	pptAddShapeCmd.Flags().String("text", "", "Text content")
	pptAddShapeCmd.Flags().Int("font-size", 0, "Font size in hundredths of a point (default 1800)")
	pptAddShapeCmd.Flags().Bool("bold", false, "Bold text")
	pptAddShapeCmd.Flags().String("color", "", "Text color as hex RGB")
	pptAddShapeCmd.Flags().String("fill", "", "Shape fill color as hex RGB")
	pptAddShapeCmd.Flags().String("line", "", "Shape outline color as hex RGB")
	pptAddShapeCmd.Flags().String("output", "", "Output .pptx path (defaults to input)")
	pptCmd.AddCommand(pptAddShapeCmd)

	markWrite(pptReplaceCmd)
	markWrite(pptImagesCmd)
	markWrite(pptCreateCmd)
	markWrite(pptAddSlideCmd)
	markWrite(pptSetContentCmd)
	markWrite(pptSetNotesCmd)
	markWrite(pptDeleteSlideCmd)
	markWrite(pptReorderCmd)
	markWrite(pptAddImageCmd)
	markWrite(pptBuildCmd)
	markWrite(pptSetStyleCmd)
	markWrite(pptAddShapeCmd)
}

var pptReadCmd = &cobra.Command{
	Use:   "read <FILE>",
	Short: "Read slide outlines (titles + bullets), markdown, or full text",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideOnly, _ := cmd.Flags().GetInt("slide")
		format, _ := cmd.Flags().GetString("format")
		keyword, _ := cmd.Flags().GetString("keyword")
		withNotes, _ := cmd.Flags().GetBool("with-notes")

		slides, err := ppt.ReadSlides(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if slideOnly > 0 {
			filtered := slides[:0]
			for _, s := range slides {
				if s.Index == slideOnly {
					filtered = append(filtered, s)
				}
			}
			slides = filtered
			if len(slides) == 0 {
				return emitError(fmt.Sprintf("slide %d not found", slideOnly), output.ErrNotFound, path, ExitNotFound)
			}
		}
		if keyword != "" {
			needle := strings.ToLower(keyword)
			filtered := slides[:0]
			for _, s := range slides {
				hay := strings.ToLower(s.Title + " " + s.Text + " " + s.Notes)
				if strings.Contains(hay, needle) {
					filtered = append(filtered, s)
				}
			}
			slides = filtered
		}
		if !withNotes {
			for i := range slides {
				slides[i].Notes = ""
			}
		}

		switch format {
		case "markdown":
			md := slidesAsMarkdown(slides, withNotes)
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "markdown": md})
				return nil
			}
			fmt.Println(md)
			return nil
		case "text":
			var sb strings.Builder
			for _, s := range slides {
				sb.WriteString("[slide ")
				sb.WriteString(fmt.Sprintf("%d", s.Index))
				sb.WriteString("] ")
				sb.WriteString(s.Title)
				sb.WriteString("\n")
				sb.WriteString(s.Text)
				sb.WriteString("\n\n")
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "text": sb.String()})
				return nil
			}
			fmt.Println(sb.String())
			return nil
		case "outline", "":
			if jsonMode {
				flat := make([]output.FlatSlide, 0, len(slides))
				for _, s := range slides {
					flat = append(flat, output.FlatSlide{
						Index: s.Index, Title: s.Title, Bullets: s.Bullets, Notes: s.Notes, Text: s.Text,
					})
				}
				output.PrintJSON(map[string]any{"file": path, "slides": flat})
				return nil
			}
			rows := make([][]string, 0, len(slides))
			for _, s := range slides {
				bullets := strings.Join(s.Bullets, " | ")
				rows = append(rows, []string{fmt.Sprintf("%d", s.Index), s.Title, bullets})
			}
			output.Table([]string{"#", "title", "bullets"}, rows)
			return nil
		default:
			return emitError("invalid --format (outline | markdown | text)", output.ErrValidation, "", ExitBadArgs)
		}
	},
}

// slidesAsMarkdown renders slides as a hierarchical Markdown outline.
// Title becomes "## ", bullets become "- ", notes become a quoted block when included.
func slidesAsMarkdown(slides []ppt.Slide, withNotes bool) string {
	var sb strings.Builder
	for _, s := range slides {
		sb.WriteString(fmt.Sprintf("## %d. %s\n\n", s.Index, s.Title))
		for _, b := range s.Bullets {
			sb.WriteString("- ")
			sb.WriteString(b)
			sb.WriteString("\n")
		}
		if withNotes && s.Notes != "" {
			sb.WriteString("\n> Notes: ")
			sb.WriteString(s.Notes)
			sb.WriteString("\n")
		}
		sb.WriteString("\n")
	}
	return sb.String()
}

var pptReplaceCmd = &cobra.Command{
	Use:   "replace <FILE>",
	Short: "Find/replace text across every slide of a .pptx (creates a new file)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		find, _ := cmd.Flags().GetString("find")
		replace, _ := cmd.Flags().GetString("replace")
		pairsRaw, _ := cmd.Flags().GetString("pairs")
		out, _ := cmd.Flags().GetString("output")

		var pairs []ppt.Replacement
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
			pairs = append(pairs, ppt.Replacement{Find: find, Replace: replace})
		}
		if len(pairs) == 0 {
			return emitError("provide --find/--replace or --pairs", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = strings.TrimSuffix(path, ".pptx") + ".replaced.pptx"
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("replace text", map[string]any{"file": path, "pairs": len(pairs), "output": out}) {
			return nil
		}

		hits, err := ppt.ReplaceText(path, out, pairs)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			payload := map[string]any{"status": "ok", "file": path, "output": out, "hits": hits}
			if hits == 0 {
				payload["warning"] = "no replacements made — text may be split across multiple runs; only text within a single <a:t> element is matched"
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

var pptCountCmd = &cobra.Command{
	Use:   "count <FILE>",
	Short: "Print just the slide count (useful for quick checks)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		n, err := ppt.SlideCount(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "slides": n})
			return nil
		}
		fmt.Println(n)
		return nil
	},
}

var pptOutlineCmd = &cobra.Command{
	Use:   "outline <FILE>",
	Short: "Lightweight title-only outline (faster than `ppt read`)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		outline, err := ppt.ReadOutline(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "outline": outline})
			return nil
		}
		rows := make([][]string, 0, len(outline))
		for _, o := range outline {
			rows = append(rows, []string{fmt.Sprintf("%d", o.Index), o.Title})
		}
		output.Table([]string{"#", "title"}, rows)
		return nil
	},
}

var pptImagesCmd = &cobra.Command{
	Use:   "images <FILE>",
	Short: "Extract every embedded image from a .pptx into --output-dir",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		outDir, _ := cmd.Flags().GetString("output-dir")
		if outDir == "" {
			return emitError("--output-dir is required", output.ErrValidation, "", ExitBadArgs)
		}
		outDir, err = resolveOutput(outDir)
		if err != nil {
			return err
		}

		if dryRunOutput("extract images", map[string]any{"file": path, "outputDir": outDir}) {
			return nil
		}

		images, err := ppt.ExtractImages(path, outDir)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "outputDir": outDir, "images": images})
			return nil
		}
		if len(images) == 0 {
			output.Gray("(no embedded images)")
			return nil
		}
		rows := make([][]string, 0, len(images))
		for _, img := range images {
			rows = append(rows, []string{img.Name, img.MediaType, fmt.Sprintf("%d", img.Bytes), img.Path})
		}
		output.Table([]string{"name", "type", "bytes", "path"}, rows)
		return nil
	},
}

var pptMetaCmd = &cobra.Command{
	Use:   "meta <FILE>",
	Short: "Show presentation metadata (title, author, slide count, etc.)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		meta, err := ppt.ReadMeta(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(meta)
			return nil
		}
		rows := [][]string{
			{"path", meta.Path},
			{"size", fmt.Sprintf("%d bytes", meta.SizeBytes)},
			{"modified", meta.Modified},
			{"title", meta.Title},
			{"author", meta.Author},
			{"subject", meta.Subject},
			{"keywords", meta.Keywords},
			{"description", meta.Description},
			{"application", meta.Application},
			{"company", meta.Company},
			{"slides", fmt.Sprintf("%d", meta.Slides)},
		}
		output.Table([]string{"field", "value"}, rows)
		return nil
	},
}
