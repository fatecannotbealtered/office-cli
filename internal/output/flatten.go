package output

import "strings"

// FlatCell is the token-efficient representation of an Excel cell.
// Only non-empty fields are included in the JSON output.
type FlatCell struct {
	Sheet   string `json:"sheet"`
	Ref     string `json:"ref"`
	Value   string `json:"value,omitempty"`
	Formula string `json:"formula,omitempty"`
	Type    string `json:"type,omitempty"`
}

// FlatSheet describes a worksheet at a glance: name, dimensions, optional preview rows.
type FlatSheet struct {
	Name      string     `json:"name"`
	Index     int        `json:"index"`
	Rows      int        `json:"rows"`
	Cols      int        `json:"cols"`
	Hidden    bool       `json:"hidden,omitempty"`
	Dimension string     `json:"dimension,omitempty"`
	Preview   [][]string `json:"preview,omitempty"`
}

// FlatPage represents one page of a PDF document.
type FlatPage struct {
	Page      int    `json:"page"`
	Text      string `json:"text,omitempty"`
	WordCount int    `json:"wordCount,omitempty"`
	Width     int    `json:"width,omitempty"`
	Height    int    `json:"height,omitempty"`
}

// FlatSlide represents one slide in a presentation.
type FlatSlide struct {
	Index   int      `json:"index"`
	Title   string   `json:"title,omitempty"`
	Bullets []string `json:"bullets,omitempty"`
	Notes   string   `json:"notes,omitempty"`
	Text    string   `json:"text,omitempty"`
}

// FlatParagraph represents one paragraph of a Word document.
type FlatParagraph struct {
	Index int    `json:"index"`
	Style string `json:"style,omitempty"`
	Text  string `json:"text"`
}

// FlatBodyElement represents one top-level item inside a Word document body —
// either a paragraph or a table — for use with --with-tables output.
type FlatBodyElement struct {
	Index int        `json:"index"`
	Type  string     `json:"type"`
	Style string     `json:"style,omitempty"`
	Text  string     `json:"text,omitempty"`
	Rows  [][]string `json:"rows,omitempty"`
}

// FlatDocMeta is the universal metadata shape for any office document.
type FlatDocMeta struct {
	Path        string   `json:"path"`
	Format      string   `json:"format"`
	SizeBytes   int64    `json:"sizeBytes"`
	Modified    string   `json:"modified,omitempty"`
	Title       string   `json:"title,omitempty"`
	Author      string   `json:"author,omitempty"`
	Subject     string   `json:"subject,omitempty"`
	Keywords    string   `json:"keywords,omitempty"`
	Description string   `json:"description,omitempty"`
	Application string   `json:"application,omitempty"`
	Created     string   `json:"created,omitempty"`
	Pages       int      `json:"pages,omitempty"`
	Slides      int      `json:"slides,omitempty"`
	Paragraphs  int      `json:"paragraphs,omitempty"`
	Sheets      []string `json:"sheets,omitempty"`
}

// FilterMap reduces m to the keys named in fieldNames (case-insensitive).
// When fieldNames is empty, m is returned unchanged. Used by --fields flag.
func FilterMap(m map[string]any, fieldNames []string) map[string]any {
	if len(fieldNames) == 0 {
		return m
	}
	lower := make(map[string]any, len(m))
	for k, v := range m {
		lower[strings.ToLower(k)] = v
	}
	result := make(map[string]any, len(fieldNames))
	for _, name := range fieldNames {
		key := strings.TrimSpace(strings.ToLower(name))
		if v, ok := lower[key]; ok {
			result[key] = v
		}
	}
	return result
}
