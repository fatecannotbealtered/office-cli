package cmd

import (
	"fmt"
	"os"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
	"github.com/fatecannotbealtered/office-cli/internal/engine/excel"
	"github.com/fatecannotbealtered/office-cli/internal/engine/pdf"
	"github.com/fatecannotbealtered/office-cli/internal/engine/ppt"
	"github.com/fatecannotbealtered/office-cli/internal/engine/word"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// extractTextCmd is the format-agnostic "give me the text content" entry point.
//
// AI Agents often receive a file of unknown format (e.g. user-uploaded
// "report.???"); this command sniffs the extension and dispatches to the right
// engine, so the Agent does not have to write four code paths.
var extractTextCmd = &cobra.Command{
	Use:   "extract-text <FILE>",
	Short: "Extract plain text from any supported format (auto-detects docx / xlsx / pptx / pdf)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx", "xlsx", "pptx", "pdf")
		if err != nil {
			return err
		}
		format := common.DetectFormat(path)

		text, err := extractText(path, format)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}

		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "format": format, "text": text})
			return nil
		}
		fmt.Println(text)
		return nil
	},
}

// extractText is the dispatcher. Each engine returns its best plain-text view:
//   - docx: paragraphs separated by newlines (no Markdown decoration)
//   - xlsx: each sheet rendered as a Markdown-style block, rows tab-separated
//   - pptx: slides separated by a "[slide N] title\n..." header
//   - pdf:  per-page text concatenated with form-feed page breaks
func extractText(path, format string) (string, error) {
	switch format {
	case "docx":
		paragraphs, err := word.ReadParagraphs(path)
		if err != nil {
			return "", err
		}
		var sb strings.Builder
		for _, p := range paragraphs {
			sb.WriteString(p.Text)
			sb.WriteString("\n")
		}
		return sb.String(), nil

	case "xlsx":
		sheets, err := excel.ListSheets(path)
		if err != nil {
			return "", err
		}
		var sb strings.Builder
		for _, s := range sheets {
			sb.WriteString("## ")
			sb.WriteString(s.Name)
			sb.WriteString("\n")
			res, err := excel.Read(path, excel.ReadOptions{Sheet: s.Name})
			if err != nil {
				continue
			}
			for _, row := range res.Rows {
				sb.WriteString(strings.Join(row, "\t"))
				sb.WriteString("\n")
			}
			sb.WriteString("\n")
		}
		return sb.String(), nil

	case "pptx":
		slides, err := ppt.ReadSlides(path)
		if err != nil {
			return "", err
		}
		var sb strings.Builder
		for _, s := range slides {
			sb.WriteString(fmt.Sprintf("[slide %d] %s\n", s.Index, s.Title))
			sb.WriteString(s.Text)
			sb.WriteString("\n\n")
		}
		return sb.String(), nil

	case "pdf":
		return pdf.AllText(path)
	}
	return "", fmt.Errorf("unsupported format: %s", format)
}

// metaCmd is the format-agnostic metadata reader. Output is a unified
// FlatDocMeta so AI Agents can ingest "any" file and inspect its summary fields
// (title, author, page/slide count, sheet list) without branching by format.
var metaCmd = &cobra.Command{
	Use:   "meta <FILE>",
	Short: "Read metadata from any supported format (auto-detects docx / xlsx / pptx / pdf)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx", "xlsx", "pptx", "pdf")
		if err != nil {
			return err
		}
		format := common.DetectFormat(path)
		stat, _ := os.Stat(path)

		out := output.FlatDocMeta{
			Path:   path,
			Format: format,
		}
		if stat != nil {
			out.SizeBytes = stat.Size()
			out.Modified = stat.ModTime().Format("2006-01-02T15:04:05Z07:00")
		}

		switch format {
		case "docx":
			m, err := word.ReadMeta(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			out.Title, out.Author, out.Subject = m.Title, m.Author, m.Subject
			out.Keywords, out.Description, out.Application = m.Keywords, m.Description, m.Application
			out.Created = m.Created
			out.Pages = m.Pages
			out.Paragraphs = m.Paragraphs
		case "pptx":
			m, err := ppt.ReadMeta(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			out.Title, out.Author, out.Subject = m.Title, m.Author, m.Subject
			out.Keywords, out.Description, out.Application = m.Keywords, m.Description, m.Application
			out.Slides = m.Slides
		case "pdf":
			info, err := pdf.FullInfo(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			out.Title = stringFromAny(info["title"])
			out.Author = stringFromAny(info["author"])
			out.Subject = stringFromAny(info["subject"])
			out.Application = stringFromAny(info["producer"])
			out.Created = stringFromAny(info["creationDate"])
			if v, ok := info["pageCount"].(int); ok {
				out.Pages = v
			}
		case "xlsx":
			sheets, err := excel.ListSheets(path)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			names := make([]string, 0, len(sheets))
			for _, s := range sheets {
				names = append(names, s.Name)
			}
			out.Sheets = names
		}

		if jsonMode {
			output.PrintJSON(out)
			return nil
		}
		rows := [][]string{
			{"path", out.Path},
			{"format", out.Format},
			{"sizeBytes", fmt.Sprintf("%d", out.SizeBytes)},
			{"modified", out.Modified},
			{"title", out.Title},
			{"author", out.Author},
			{"subject", out.Subject},
			{"application", out.Application},
			{"created", out.Created},
		}
		if out.Pages > 0 {
			rows = append(rows, []string{"pages", fmt.Sprintf("%d", out.Pages)})
		}
		if out.Slides > 0 {
			rows = append(rows, []string{"slides", fmt.Sprintf("%d", out.Slides)})
		}
		if out.Paragraphs > 0 {
			rows = append(rows, []string{"paragraphs", fmt.Sprintf("%d", out.Paragraphs)})
		}
		if len(out.Sheets) > 0 {
			rows = append(rows, []string{"sheets", strings.Join(out.Sheets, ", ")})
		}
		output.Table([]string{"field", "value"}, rows)
		return nil
	},
}

// infoCmd is a single-line "what is this file" probe. Faster than meta because
// it only reports format + size + the most critical counter (pages/slides/sheets).
var infoCmd = &cobra.Command{
	Use:   "info <FILE>",
	Short: "Quickly identify the format and primary counter of any supported file",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "docx", "xlsx", "pptx", "pdf")
		if err != nil {
			return err
		}
		format := common.DetectFormat(path)
		stat, _ := os.Stat(path)

		summary := map[string]any{"path": path, "format": format}
		if stat != nil {
			summary["sizeBytes"] = stat.Size()
		} else {
			return emitError("cannot stat file: "+path, output.ErrEngine, path, ExitEngine)
		}

		switch format {
		case "docx":
			paragraphs, err := word.ReadParagraphs(path)
			if err == nil {
				summary["paragraphs"] = len(paragraphs)
				summary["headings"] = len(word.ExtractHeadings(paragraphs))
			}
		case "xlsx":
			sheets, err := excel.ListSheets(path)
			if err == nil {
				summary["sheets"] = len(sheets)
			}
		case "pptx":
			n, err := ppt.SlideCount(path)
			if err == nil {
				summary["slides"] = n
			}
		case "pdf":
			n, err := pdf.PageCount(path)
			if err == nil {
				summary["pages"] = n
			}
		}

		if jsonMode {
			output.PrintJSON(summary)
			return nil
		}
		fmt.Printf("%s  %s  %d bytes", path, format, summary["sizeBytes"])
		for _, key := range []string{"pages", "slides", "sheets", "paragraphs"} {
			if v, ok := summary[key]; ok {
				fmt.Printf("  %s=%v", key, v)
			}
		}
		fmt.Println()
		return nil
	},
}

// stringFromAny converts an interface value to a string, returning "" for nil.
func stringFromAny(v any) string {
	if v == nil {
		return ""
	}
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprintf("%v", v)
}

func init() {
	rootCmd.AddCommand(extractTextCmd)
	rootCmd.AddCommand(metaCmd)
	rootCmd.AddCommand(infoCmd)
}
