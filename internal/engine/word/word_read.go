package word

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

func ReadParagraphs(path string) ([]Paragraph, error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = zc.Close() }()

	data, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	if err != nil {
		return nil, err
	}

	var doc xmlBody
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, fmt.Errorf("parsing document.xml: %w", err)
	}

	out := make([]Paragraph, 0, len(doc.Body.Paragraphs))
	for i, p := range doc.Body.Paragraphs {
		var sb strings.Builder
		for _, r := range p.Runs {
			for _, t := range r.Text {
				sb.WriteString(t.Value)
			}
		}
		text := sb.String()
		// Skip purely empty paragraphs unless they have a heading style; helps reduce noise.
		if text == "" && p.Pr.Style.Val == "" {
			continue
		}
		out = append(out, Paragraph{
			Index: i,
			Style: p.Pr.Style.Val,
			Text:  text,
		})
	}
	return out, nil
}

// AsMarkdown formats paragraphs as Markdown using their pStyle hints.
// Heading1..Heading6 → "#"…"######". Lists are not detected at this level
// (they require numbering.xml lookup); use ReadParagraphs and post-process if needed.
func AsMarkdown(paragraphs []Paragraph) string {
	var sb strings.Builder
	for _, p := range paragraphs {
		switch {
		case strings.HasPrefix(strings.ToLower(p.Style), "heading"):
			level := 1
			if n := strings.TrimPrefix(strings.ToLower(p.Style), "heading"); n != "" {
				if v := parseHeadingLevel(n); v > 0 && v <= 6 {
					level = v
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
	return sb.String()
}

// parseHeadingLevel converts the numeric tail of a "Heading3" style into 3.
// Returns 0 when not parseable.
func parseHeadingLevel(s string) int {
	level := 0
	for _, r := range s {
		if r < '0' || r > '9' {
			break
		}
		level = level*10 + int(r-'0')
	}
	return level
}

// Replacement is one find-replace pair.
type Replacement struct {
	Find    string `json:"find"`
	Replace string `json:"replace"`
}

// ReplaceText rewrites word/document.xml replacing every occurrence of each
// `find` string with the matching `replace` value. The file at outPath is created
// (or overwritten) with the new content; the source path is never modified.
//
// Limitation: replacement only matches text that lives entirely inside a single
// <w:t> element. Replacements spanning multiple runs are not detected.
func ReplaceText(path, outPath string, replacements []Replacement) (int, error) {
	if len(replacements) == 0 {
		return 0, errors.New("no replacements provided")
	}

	srcF, err := os.Open(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = srcF.Close() }()
	stat, err := srcF.Stat()
	if err != nil {
		return 0, err
	}
	zr, err := zip.NewReader(srcF, stat.Size())
	if err != nil {
		return 0, err
	}

	doc, err := common.ReadEntry(zr, documentXMLPath)
	if err != nil {
		return 0, err
	}

	hits := 0
	current := doc
	for _, r := range replacements {
		if r.Find == "" {
			continue
		}
		hits += strings.Count(string(current), r.Find)
		current = common.ReplaceInBytes(current, r.Find, r.Replace)
	}

	if err := common.RewriteEntries(path, outPath, map[string][]byte{
		documentXMLPath: current,
	}); err != nil {
		return 0, err
	}
	return hits, nil
}

// ReadMeta combines core.xml, app.xml and filesystem stats. Best-effort: missing
// fields are returned as empty strings / zero values rather than as errors.
func ReadMeta(path string) (Meta, error) {
	stat, err := os.Stat(path)
	if err != nil {
		return Meta{}, err
	}
	zc, err := common.OpenReader(path)
	if err != nil {
		return Meta{}, err
	}
	defer func() { _ = zc.Close() }()

	core := common.ReadCoreProps(&zc.Reader)
	app := common.ReadAppProps(&zc.Reader)

	paragraphs, _ := ReadParagraphs(path)

	return Meta{
		Path:        path,
		SizeBytes:   stat.Size(),
		Modified:    stat.ModTime().Format("2006-01-02T15:04:05Z07:00"),
		Title:       strings.TrimSpace(core.Title),
		Author:      strings.TrimSpace(core.Creator),
		Subject:     strings.TrimSpace(core.Subject),
		Keywords:    strings.TrimSpace(core.Keywords),
		Description: strings.TrimSpace(core.Description),
		Application: strings.TrimSpace(app.Application),
		Created:     strings.TrimSpace(core.Created),
		Pages:       app.Pages,
		Words:       app.Words,
		Paragraphs:  len(paragraphs),
	}, nil
}

// CountWords counts words across paragraphs (whitespace tokens). Used as a fallback
// when app.xml does not embed a word count.
func CountWords(paragraphs []Paragraph) int {
	total := 0
	for _, p := range paragraphs {
		total += len(strings.Fields(p.Text))
	}
	return total
}

// Stats summarizes the document size in human-meaningful units.
type Stats struct {
	Paragraphs int `json:"paragraphs"`
	Headings   int `json:"headings"`
	Words      int `json:"words"`
	Characters int `json:"characters"`
	Lines      int `json:"lines"`
}

// ComputeStats walks the paragraphs once and returns aggregate counts.
// Useful when AI Agents want a one-shot "how big is this document".
func ComputeStats(paragraphs []Paragraph) Stats {
	s := Stats{Paragraphs: len(paragraphs)}
	for _, p := range paragraphs {
		if strings.HasPrefix(strings.ToLower(p.Style), "heading") || strings.EqualFold(p.Style, "Title") {
			s.Headings++
		}
		s.Words += len(strings.Fields(p.Text))
		s.Characters += len([]rune(p.Text))
		if p.Text != "" {
			s.Lines++
		}
	}
	return s
}

// Heading is one entry in a document's heading outline.
type Heading struct {
	Index int    `json:"index"`
	Level int    `json:"level"`
	Text  string `json:"text"`
}

// ExtractHeadings returns just the headings (Heading1..Heading6 + Title) from the
// document. Each entry includes its 1-based heading level. Title becomes level 1.
func ExtractHeadings(paragraphs []Paragraph) []Heading {
	var headings []Heading
	for _, p := range paragraphs {
		style := strings.ToLower(p.Style)
		switch {
		case strings.EqualFold(p.Style, "Title"):
			headings = append(headings, Heading{Index: p.Index, Level: 1, Text: p.Text})
		case strings.HasPrefix(style, "heading"):
			level := parseHeadingLevel(strings.TrimPrefix(style, "heading"))
			if level <= 0 {
				level = 1
			}
			if level > 6 {
				level = 6
			}
			headings = append(headings, Heading{Index: p.Index, Level: level, Text: p.Text})
		}
	}
	return headings
}

// ExtractedImage represents one media file pulled out of a docx.
type ExtractedImage struct {
	Name      string `json:"name"`
	Path      string `json:"path"`
	Bytes     int64  `json:"bytes"`
	MediaType string `json:"mediaType,omitempty"`
}

// ExtractImages dumps every file under word/media/ to outDir. The original file
// names are preserved (image1.png, image2.jpg, ...). Returns metadata about
// each file written.
func ExtractImages(path, outDir string) ([]ExtractedImage, error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = zc.Close() }()

	if err := os.MkdirAll(outDir, 0755); err != nil {
		return nil, err
	}

	var out []ExtractedImage
	for _, f := range zc.File {
		if !strings.HasPrefix(f.Name, "word/media/") {
			continue
		}
		base := strings.TrimPrefix(f.Name, "word/media/")
		if base == "" {
			continue
		}
		dst := filepath.Join(outDir, base)
		if err := common.WriteZipEntryToFile(f, dst); err != nil {
			return nil, err
		}
		out = append(out, ExtractedImage{
			Name:      base,
			Path:      dst,
			Bytes:     int64(f.UncompressedSize64),
			MediaType: common.GuessMediaType(base),
		})
	}
	return out, nil
}
