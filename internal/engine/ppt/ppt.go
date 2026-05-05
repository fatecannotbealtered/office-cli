// Package ppt reads and edits .pptx (Office Open XML PowerPoint) files by
// parsing the underlying ZIP+XML structure directly. Same trade-offs as the
// word package: small dependency surface, predictable behavior, deliberately
// simple at the cost of not understanding every layout-driven nicety.
//
// Source files:
//   - ppt.go      — types and XML struct definitions
//   - ppt_read.go  — ReadSlides, ReadOutline, ReadMeta, ReplaceText, ExtractImages, SlideCount
//   - ppt_write.go — Create, AddSlide, SetSlideContent, SetNotes, DeleteSlide, Reorder, Build
//   - ppt_shape.go — ReadSlideLayout, SetShapeStyle, AddShape, shape tree parsing
package ppt

import (
	"encoding/xml"
)

// Slide is one slide's extracted content.
type Slide struct {
	Index   int      `json:"index"`
	File    string   `json:"file"`
	Title   string   `json:"title,omitempty"`
	Bullets []string `json:"bullets,omitempty"`
	Notes   string   `json:"notes,omitempty"`
	Text    string   `json:"text"`
}

// Meta is the consolidated metadata for a presentation.
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
	Company     string `json:"company,omitempty"`
	Slides      int    `json:"slides,omitempty"`
}

// ---------------------------------------------------------------------------
// Shape layout types — returned by ReadSlideLayout
// ---------------------------------------------------------------------------

// ShapeInfo describes one shape on a slide (position, size, text, styling).
type ShapeInfo struct {
	Index      int             `json:"index"`
	Type       string          `json:"type"` // "sp" | "pic" | "grpSp" | "cxnSp"
	Name       string          `json:"name,omitempty"`
	Ph         string          `json:"ph,omitempty"` // placeholder type: title, ctrTitle, body, dt, sldNum, etc.
	X          int             `json:"x"`
	Y          int             `json:"y"`
	W          int             `json:"w"`
	H          int             `json:"h"`
	Text       string          `json:"text"`
	Paragraphs []ParagraphInfo `json:"paragraphs,omitempty"`
}

// ParagraphInfo describes one paragraph within a shape's text body.
type ParagraphInfo struct {
	Align string    `json:"align,omitempty"` // l, r, ctr, just
	Runs  []RunInfo `json:"runs"`
}

// RunInfo describes one text run within a paragraph.
type RunInfo struct {
	Text      string `json:"text"`
	FontSize  int    `json:"fontSize,omitempty"` // hundredths of a point (3600 = 36pt)
	Bold      bool   `json:"bold,omitempty"`
	Italic    bool   `json:"italic,omitempty"`
	Color     string `json:"color,omitempty"` // hex RGB without # (e.g. "FF0000")
	Underline bool   `json:"underline,omitempty"`
}

// StyleOptions specifies which text styling properties to modify.
// Nil fields are left unchanged.
type StyleOptions struct {
	FontSize  *int
	Bold      *bool
	Italic    *bool
	Underline *bool
	Color     *string
	Align     *string
}

// ShapeSpec describes a new shape to insert via AddShape.
type ShapeSpec struct {
	Type     string `json:"type"` // "text-box"|"rect"|"ellipse"|"line"|"arrow"
	X        int    `json:"x"`
	Y        int    `json:"y"`
	W        int    `json:"w"`
	H        int    `json:"h"`
	Text     string `json:"text,omitempty"`
	FontSize int    `json:"fontSize,omitempty"` // hundredths of a point (2400 = 24pt)
	Bold     bool   `json:"bold,omitempty"`
	Color    string `json:"color,omitempty"` // text color hex
	Fill     string `json:"fill,omitempty"`  // shape fill hex
	Line     string `json:"line,omitempty"`  // outline color hex
}

// xmlSlide mirrors enough of one slideN.xml to recover its text.
type xmlSlide struct {
	XMLName xml.Name `xml:"sld"`
	CSld    struct {
		SpTree struct {
			Shapes []xmlShape `xml:"sp"`
			GFrame []struct {
				Graphic struct {
					GraphicData struct {
						Table struct {
							TR []struct {
								TC []struct {
									TxBody xmlTxBody `xml:"txBody"`
								} `xml:"tc"`
							} `xml:"tr"`
						} `xml:"tbl"`
					} `xml:"graphicData"`
				} `xml:"graphic"`
			} `xml:"graphicFrame"`
		} `xml:"spTree"`
	} `xml:"cSld"`
}

type xmlShape struct {
	NvSpPr struct {
		NvPr struct {
			Ph struct {
				Type string `xml:"type,attr"`
				Idx  string `xml:"idx,attr"`
			} `xml:"ph"`
		} `xml:"nvPr"`
	} `xml:"nvSpPr"`
	TxBody xmlTxBody `xml:"txBody"`
}

type xmlTxBody struct {
	Paragraphs []xmlPPTParagraph `xml:"p"`
}

type xmlPPTParagraph struct {
	Runs []xmlPPTRun `xml:"r"`
}

type xmlPPTRun struct {
	Text string `xml:"t"`
}

// slideEntry pairs a slide file number with its zip entry name.
type slideEntry struct {
	fileNum   int
	entryName string
}

// ---------------------------------------------------------------------------
// Shared XML types for presentation.xml and .rels parsing
// ---------------------------------------------------------------------------

// xmlRelEntry is one <Relationship> parsed from a .rels file.
type xmlRelEntry struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// xmlRels is the root of a .rels file.
type xmlRels struct {
	Entries []xmlRelEntry `xml:"Relationship"`
}

// slideRelMap maps slide file number to the rId used in presentation.xml.
type slideRelMap struct {
	slideToRId map[int]string
	rIDToSlide map[string]int
	maxSldID   int
	maxRIdNum  int
}

// sldIDEntry holds one parsed <p:sldId> entry.
type sldIDEntry struct {
	SldID int
	RId   string
}

// Replacement is one find-replace pair, mirroring word.Replacement.
type Replacement struct {
	Find    string `json:"find"`
	Replace string `json:"replace"`
}
