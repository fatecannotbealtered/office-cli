package cmd

import (
	"fmt"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/word"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var wordCmd = &cobra.Command{
	Use:   "word",
	Short: "Read paragraphs, replace text and inspect metadata of .docx files",
}

func init() {
	rootCmd.AddCommand(wordCmd)

	wordReadCmd.Flags().String("format", "paragraphs", "Output format: paragraphs | markdown | text")
	wordReadCmd.Flags().Int("limit", 0, "Max paragraphs to return (0 = all)")
	wordReadCmd.Flags().String("keyword", "", "Only return paragraphs containing this keyword (case-insensitive)")
	wordReadCmd.Flags().Bool("with-tables", false, "Include tables in the output (paragraphs + tables in document order)")
	wordCmd.AddCommand(wordReadCmd)

	wordCmd.AddCommand(wordMetaCmd)
	wordCmd.AddCommand(wordStatsCmd)
	wordCmd.AddCommand(wordHeadingsCmd)

	wordImagesCmd.Flags().String("output-dir", "", "Directory to write extracted images (required)")
	wordCmd.AddCommand(wordImagesCmd)

	markWrite(wordImagesCmd)
}

// ---------------------------------------------------------------------------
// Read commands
// ---------------------------------------------------------------------------

var wordReadCmd = &cobra.Command{
	Use:   "read <FILE>",
	Short: "Read paragraphs from a .docx (paragraphs | markdown | text)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		format, _ := cmd.Flags().GetString("format")
		limit, _ := cmd.Flags().GetInt("limit")
		keyword, _ := cmd.Flags().GetString("keyword")
		withTables, _ := cmd.Flags().GetBool("with-tables")

		if withTables {
			return wordReadWithTables(path, format, keyword, limit)
		}

		paragraphs, err := word.ReadParagraphs(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if keyword != "" {
			needle := strings.ToLower(keyword)
			filtered := paragraphs[:0]
			for _, p := range paragraphs {
				if strings.Contains(strings.ToLower(p.Text), needle) {
					filtered = append(filtered, p)
				}
			}
			paragraphs = filtered
		}
		if limit > 0 && len(paragraphs) > limit {
			paragraphs = paragraphs[:limit]
		}

		switch format {
		case "markdown":
			md := word.AsMarkdown(paragraphs)
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "markdown": md})
				return nil
			}
			fmt.Println(md)
			return nil
		case "text":
			var sb strings.Builder
			for _, p := range paragraphs {
				sb.WriteString(p.Text)
				sb.WriteString("\n")
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "text": sb.String()})
				return nil
			}
			fmt.Println(sb.String())
			return nil
		case "paragraphs", "":
			if jsonMode {
				flat := make([]output.FlatParagraph, 0, len(paragraphs))
				for _, p := range paragraphs {
					flat = append(flat, output.FlatParagraph{Index: p.Index, Style: p.Style, Text: p.Text})
				}
				output.PrintJSON(map[string]any{"file": path, "paragraphs": flat})
				return nil
			}
			rows := make([][]string, 0, len(paragraphs))
			for _, p := range paragraphs {
				rows = append(rows, []string{fmt.Sprintf("%d", p.Index), p.Style, p.Text})
			}
			output.Table([]string{"#", "style", "text"}, rows)
			return nil
		default:
			return emitError("invalid --format (paragraphs | markdown | text)", output.ErrValidation, "", ExitBadArgs)
		}
	},
}

var wordStatsCmd = &cobra.Command{
	Use:   "stats <FILE>",
	Short: "Compute paragraph / heading / word / character / line counts",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		paragraphs, err := word.ReadParagraphs(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		stats := word.ComputeStats(paragraphs)
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "stats": stats})
			return nil
		}
		rows := [][]string{
			{"paragraphs", fmt.Sprintf("%d", stats.Paragraphs)},
			{"headings", fmt.Sprintf("%d", stats.Headings)},
			{"words", fmt.Sprintf("%d", stats.Words)},
			{"characters", fmt.Sprintf("%d", stats.Characters)},
			{"lines", fmt.Sprintf("%d", stats.Lines)},
		}
		output.Table([]string{"metric", "value"}, rows)
		return nil
	},
}

var wordHeadingsCmd = &cobra.Command{
	Use:   "headings <FILE>",
	Short: "Extract just the heading outline (Title / Heading1..6)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		paragraphs, err := word.ReadParagraphs(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		headings := word.ExtractHeadings(paragraphs)
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "headings": headings})
			return nil
		}
		rows := make([][]string, 0, len(headings))
		for _, h := range headings {
			rows = append(rows, []string{
				fmt.Sprintf("%d", h.Level),
				strings.Repeat("  ", h.Level-1) + h.Text,
			})
		}
		output.Table([]string{"level", "text"}, rows)
		return nil
	},
}

var wordImagesCmd = &cobra.Command{
	Use:   "images <FILE>",
	Short: "Extract every embedded image from a .docx into --output-dir",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
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

		images, err := word.ExtractImages(path, outDir)
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

var wordMetaCmd = &cobra.Command{
	Use:   "meta <FILE>",
	Short: "Show document metadata (title, author, word count, etc.)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx")
		if err != nil {
			return err
		}
		meta, err := word.ReadMeta(path)
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
			{"created", meta.Created},
			{"pages", fmt.Sprintf("%d", meta.Pages)},
			{"words", fmt.Sprintf("%d", meta.Words)},
			{"paragraphs", fmt.Sprintf("%d", meta.Paragraphs)},
		}
		output.Table([]string{"field", "value"}, rows)
		return nil
	},
}

// ---------------------------------------------------------------------------
// wordReadWithTables handles --with-tables mode for word read.
// ---------------------------------------------------------------------------

func wordReadWithTables(path, format, keyword string, limit int) error {
	elements, err := word.ReadBodyElements(path)
	if err != nil {
		return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
	}

	if keyword != "" {
		needle := strings.ToLower(keyword)
		filtered := elements[:0]
		for _, e := range elements {
			match := false
			if e.Paragraph != nil {
				match = strings.Contains(strings.ToLower(e.Paragraph.Text), needle)
			}
			if e.Table != nil {
				for _, row := range e.Table.Rows {
					for _, cell := range row {
						if strings.Contains(strings.ToLower(cell), needle) {
							match = true
							break
						}
					}
					if match {
						break
					}
				}
			}
			if match {
				filtered = append(filtered, e)
			}
		}
		elements = filtered
	}
	if limit > 0 && len(elements) > limit {
		elements = elements[:limit]
	}

	switch format {
	case "markdown":
		var sb strings.Builder
		for _, e := range elements {
			if e.Paragraph != nil {
				p := e.Paragraph
				switch {
				case strings.HasPrefix(strings.ToLower(p.Style), "heading"):
					level := 1
					if n := strings.TrimPrefix(strings.ToLower(p.Style), "heading"); n != "" {
						if v, err := fmt.Sscanf(n, "%d", &level); v == 0 || err != nil {
							level = 1
						}
					}
					sb.WriteString(strings.Repeat("#", level))
					sb.WriteString(" ")
					sb.WriteString(p.Text)
					sb.WriteString("\n\n")
				case strings.EqualFold(p.Style, "Title"):
					sb.WriteString("# ")
					sb.WriteString(p.Text)
					sb.WriteString("\n\n")
				default:
					sb.WriteString(p.Text)
					sb.WriteString("\n\n")
				}
			}
			if e.Table != nil && len(e.Table.Rows) > 0 {
				rows := e.Table.Rows
				for i, row := range rows {
					sb.WriteString("| ")
					sb.WriteString(strings.Join(row, " | "))
					sb.WriteString(" |\n")
					if i == 0 {
						sb.WriteString("| ")
						sb.WriteString(strings.Repeat("--- | ", len(row)))
						sb.WriteString("\n")
					}
				}
				sb.WriteString("\n")
			}
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "markdown": sb.String()})
			return nil
		}
		fmt.Println(sb.String())
		return nil
	case "text":
		var sb strings.Builder
		for _, e := range elements {
			if e.Paragraph != nil {
				sb.WriteString(e.Paragraph.Text)
				sb.WriteString("\n")
			}
			if e.Table != nil {
				for _, row := range e.Table.Rows {
					sb.WriteString(strings.Join(row, "\t"))
					sb.WriteString("\n")
				}
			}
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "text": sb.String()})
			return nil
		}
		fmt.Println(sb.String())
		return nil
	case "paragraphs", "":
		if jsonMode {
			flat := make([]output.FlatBodyElement, 0, len(elements))
			for _, e := range elements {
				fe := output.FlatBodyElement{Index: e.Index, Type: e.Type}
				if e.Paragraph != nil {
					fe.Style = e.Paragraph.Style
					fe.Text = e.Paragraph.Text
				}
				if e.Table != nil {
					fe.Rows = e.Table.Rows
				}
				flat = append(flat, fe)
			}
			output.PrintJSON(map[string]any{"file": path, "elements": flat})
			return nil
		}
		rows := make([][]string, 0, len(elements))
		for _, e := range elements {
			if e.Paragraph != nil {
				rows = append(rows, []string{fmt.Sprintf("%d", e.Index), "paragraph", e.Paragraph.Style, e.Paragraph.Text})
			}
			if e.Table != nil {
				summary := fmt.Sprintf("%dx%d table", len(e.Table.Rows), len(e.Table.Rows[0]))
				rows = append(rows, []string{fmt.Sprintf("%d", e.Index), "table", "", summary})
			}
		}
		output.Table([]string{"#", "type", "style", "content"}, rows)
		return nil
	default:
		return emitError("invalid --format (paragraphs | markdown | text)", output.ErrValidation, "", ExitBadArgs)
	}
}
