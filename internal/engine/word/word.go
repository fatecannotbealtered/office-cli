// Package word reads and edits .docx (Office Open XML Word) files by parsing
// the underlying ZIP+XML structure directly. We deliberately avoid heavy
// third-party libraries to keep the binary small and the dependency surface
// auditable.
//
// Source files:
//   - word.go       — types, constants, and XML struct definitions
//   - word_read.go   — ReadParagraphs, ReadMeta, AsMarkdown, ExtractImages, SearchBodyElements
//   - word_write.go  — Create, AddParagraph, AddTable, AddImage, Merge, body element operations
//   - word_style.go  — StyleParagraph, StyleTable, run/paragraph/cell formatting
package word

import (
	"encoding/xml"
)

// Default font configuration — mirrors the standard Word 2016+ factory defaults.
// Used when creating new documents and when no explicit --font-family is given.
const (
	defaultFontLatin      = "Calibri"
	defaultFontEastAsia   = "宋体"
	defaultFontComplex    = "Times New Roman"
	defaultFontSizeHalfPt = "22" // 11pt in half-points
)

// Paragraph is one logical paragraph extracted from word/document.xml.
type Paragraph struct {
	Index int    `json:"index"`
	Style string `json:"style,omitempty"`
	Text  string `json:"text"`
}

// Meta combines CoreProps + AppProps + filesystem stats into a single struct.
type Meta struct {
	Path        string `json:"path"`
	SizeBytes   int64  `json:"sizeBytes"`
	Modified    string `json:"modified,omitempty"`
	Title       string `json:"title,omitempty"`
	Author      string `json:"author,omitempty"`
	Subject     string `json:"subject,omitempty"`
	Keywords    string `json:"keywords,omitempty"`
	Description string `json:"description,omitempty"`
	Application string `json:"application,omitempty"`
	Created     string `json:"created,omitempty"`
	Pages       int    `json:"pages,omitempty"`
	Words       int    `json:"words,omitempty"`
	Paragraphs  int    `json:"paragraphs,omitempty"`
}

// xmlBody mirrors enough of word/document.xml to extract paragraphs and styles.
//
// The XML is namespaced (xmlns:w="..."); we use the "Local" matching mode in
// encoding/xml by tagging fields without namespaces — Go's stdlib accepts that
// when LocalName matches.
type xmlBody struct {
	XMLName xml.Name `xml:"document"`
	Body    struct {
		Paragraphs []xmlParagraph `xml:"p"`
	} `xml:"body"`
}

type xmlParagraph struct {
	Pr struct {
		Style struct {
			Val string `xml:"val,attr"`
		} `xml:"pStyle"`
	} `xml:"pPr"`
	Runs []xmlRun `xml:"r"`
}

type xmlRun struct {
	Text []xmlText `xml:"t"`
}

type xmlText struct {
	Value string `xml:",chardata"`
}

const documentXMLPath = "word/document.xml"
