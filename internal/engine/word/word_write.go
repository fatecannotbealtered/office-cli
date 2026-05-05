package word

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

func Create(path, title, author string) error {
	if _, err := os.Stat(path); err == nil {
		return fmt.Errorf("file already exists: %s", path)
	}
	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`,
		"word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
		"word/styles.xml": fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>
      <w:sz w:val="%s"/>
      <w:szCs w:val="%s"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/></w:style>
</w:styles>`,
			defaultFontLatin, defaultFontLatin,
			defaultFontEastAsia, defaultFontComplex,
			defaultFontSizeHalfPt, defaultFontSizeHalfPt),
		"word/document.xml": fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body></w:body>
</w:document>`),
		"docProps/core.xml": fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>%s</dc:title>
  <dc:creator>%s</dc:creator>
</cp:coreProperties>`, xmlEscape(title), xmlEscape(author)),
	}
	return common.WriteNewZip(path, files)
}

// xmlEscape escapes XML special characters.
func xmlEscape(s string) string {
	var b strings.Builder
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

// injectBodyXML reads document.xml from path, injects xmlContent just before </w:body>,
// and writes the result to outPath using RewriteEntries.
func injectBodyXML(path, outPath, xmlContent string) error {
	srcF, err := os.Open(path)
	if err != nil {
		return err
	}
	defer func() { _ = srcF.Close() }()
	stat, err := srcF.Stat()
	if err != nil {
		return err
	}
	zr, err := zip.NewReader(srcF, stat.Size())
	if err != nil {
		return err
	}
	doc, err := common.ReadEntry(zr, documentXMLPath)
	if err != nil {
		return err
	}

	bodyClose := "</w:body>"
	idx := bytes.LastIndex(doc, []byte(bodyClose))
	if idx < 0 {
		return fmt.Errorf("invalid document.xml: missing %s", bodyClose)
	}

	var buf bytes.Buffer
	buf.Write(doc[:idx])
	buf.WriteString(xmlContent)
	buf.WriteString("\n")
	buf.Write(doc[idx:])

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		documentXMLPath: buf.Bytes(),
	})
}

// AddParagraph appends a paragraph with optional style to the document.
// style can be "Normal", "Title", "Heading1"-"Heading6", or "" for no style.
func AddParagraph(path, outPath, text, style string) error {
	return injectBodyXML(path, outPath, ParagraphXML(Paragraph{Style: style, Text: text}))
}

// AddHeading is a convenience wrapper for AddParagraph with a heading style.
func AddHeading(path, outPath, text string, level int) error {
	if level < 1 || level > 6 {
		return fmt.Errorf("heading level must be 1-6, got %d", level)
	}
	return AddParagraph(path, outPath, text, fmt.Sprintf("Heading%d", level))
}

// AddTable appends a table to the document. rows is a list of rows, each row is
// a list of cell strings. The first row is rendered as a header (bold).
func AddTable(path, outPath string, rows [][]string) error {
	if len(rows) == 0 {
		return errors.New("no rows provided")
	}
	return injectBodyXML(path, outPath, TableXML(rows))
}

// AddImage inserts an inline image into the document. widthPt and heightPt
// specify the image dimensions in points (72pt = 1 inch). If 0, defaults to 200x200.
func AddImage(path, outPath, imagePath string, widthPt, heightPt float64) error {
	if widthPt <= 0 {
		widthPt = 200
	}
	if heightPt <= 0 {
		heightPt = 200
	}

	// Read image data
	imgData, err := os.ReadFile(imagePath)
	if err != nil {
		return fmt.Errorf("reading image: %w", err)
	}

	// Determine content type and extension
	ext := strings.ToLower(filepath.Ext(imagePath))
	ct := common.GuessMediaType(filepath.Base(imagePath))
	if ct == "" {
		ct = "image/png"
	}

	// We need a unique relationship ID. Count existing media files in the zip.
	srcF, err := os.Open(path)
	if err != nil {
		return err
	}
	stat, err := srcF.Stat()
	if err != nil {
		_ = srcF.Close()
		return err
	}
	zr, err := zip.NewReader(srcF, stat.Size())
	if err != nil {
		_ = srcF.Close()
		return err
	}

	mediaCount := 0
	entries := map[string][]byte{}
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "word/media/") {
			mediaCount++
		}
	}

	mediaName := fmt.Sprintf("image%d%s", mediaCount+1, ext)
	rID := fmt.Sprintf("rId%d", mediaCount+10) // avoid collision with rId1 (styles)

	// Add image to word/media/
	entries["word/media/"+mediaName] = imgData

	// Update relationships
	rels, err := common.ReadEntry(zr, "word/_rels/document.xml.rels")
	if err != nil {
		// Create from scratch
		rels = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)
	}
	relsStr := string(rels)
	// Insert new relationship before closing tag
	insertPoint := strings.LastIndex(relsStr, "</Relationships>")
	if insertPoint >= 0 {
		newRel := fmt.Sprintf(`  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`,
			rID, mediaName)
		relsStr = relsStr[:insertPoint] + newRel + "\n" + relsStr[insertPoint:]
	}
	entries["word/_rels/document.xml.rels"] = []byte(relsStr)

	// Update [Content_Types].xml to include image extension if needed
	ctData, err := common.ReadEntry(zr, "[Content_Types].xml")
	if err == nil {
		ctStr := string(ctData)
		extNoDot := strings.TrimPrefix(ext, ".")
		if !strings.Contains(ctStr, fmt.Sprintf(`Extension="%s"`, extNoDot)) {
			newCT := fmt.Sprintf(`  <Default Extension="%s" ContentType="%s"/>`, extNoDot, ct)
			insertPoint := strings.LastIndex(ctStr, "</Types>")
			if insertPoint >= 0 {
				ctStr = ctStr[:insertPoint] + newCT + "\n" + ctStr[insertPoint:]
			}
		}
		entries["[Content_Types].xml"] = []byte(ctStr)
	}

	// EMU: 1 point = 12700 EMU
	emuW := int64(widthPt * 12700)
	emuH := int64(heightPt * 12700)

	// Create drawing XML
	drawing := fmt.Sprintf(`<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="%d" cy="%d"/>
    <wp:docPr id="%d" name="Image %d"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr><pic:cNvPr id="0" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`, emuW, emuH, mediaCount+1, mediaCount+1, mediaName, rID, emuW, emuH)

	para := fmt.Sprintf("<w:p><w:r>%s</w:r></w:p>", drawing)

	// Close source file before writing
	_ = srcF.Close()

	// Write media and rels first
	out := outPath
	if out == "" {
		out = path
	}
	if err := common.RewriteEntries(path, out, entries); err != nil {
		return err
	}

	// Now inject the drawing paragraph into the output file
	return injectBodyXML(out, out, para)
}

// AddPageBreak inserts a page break paragraph.
func AddPageBreak(path, outPath string) error {
	para := `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`
	return injectBodyXML(path, outPath, para)
}

// ---------------------------------------------------------------------------
// Body-element read/write: parse paragraphs AND tables in document order
// ---------------------------------------------------------------------------

// TableData holds the cell contents of one table.
type TableData struct {
	Rows [][]string `json:"rows"`
}

// BodyElement is one top-level item inside <w:body> — either a paragraph or a table.
type BodyElement struct {
	Index     int        `json:"index"`
	Type      string     `json:"type"` // "paragraph" or "table"
	Paragraph *Paragraph `json:"paragraph,omitempty"`
	Table     *TableData `json:"table,omitempty"`
}

// ReadBodyElements tokenises <w:body> and returns every child element in document
// order. Paragraphs are returned as Paragraph pointers; tables as TableData pointers.
// This is the foundation for delete / insert / update-table-cell operations.
func ReadBodyElements(path string) ([]BodyElement, error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = zc.Close() }()

	data, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	if err != nil {
		return nil, err
	}

	return parseBodyElements(data)
}

// parseBodyElements is the token-based parser that walks <w:body> children.
func parseBodyElements(data []byte) ([]BodyElement, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var elements []BodyElement
	idx := 0
	depth := 0
	inBody := false

	// State for paragraph collection
	var inParagraph bool
	var paraStyle string
	var paraText strings.Builder

	// State for table collection
	var inTable bool
	var currentRow []string
	var currentTableRows [][]string
	var cellText strings.Builder
	var inCell bool

	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			local := t.Name.Local
			switch {
			case local == "body" && depth == 2:
				inBody = true
			case inBody && local == "p":
				inParagraph = true
				paraStyle = ""
				paraText.Reset()
			case inBody && inParagraph && local == "pStyle":
				for _, a := range t.Attr {
					if a.Name.Local == "val" {
						paraStyle = a.Value
					}
				}
			case inBody && local == "tbl":
				inTable = true
				currentTableRows = nil
			case inBody && inTable && local == "tr":
				currentRow = nil
			case inBody && inTable && local == "tc":
				inCell = true
				cellText.Reset()
			case inBody && inTable && inCell && local == "p":
				// paragraph inside table cell — skip
			case inBody && local == "t":
				if inParagraph && !inTable {
					var s string
					if err := dec.DecodeElement(&s, &t); err == nil {
						paraText.WriteString(s)
					}
					depth--
					continue
				} else if inCell {
					var s string
					if err := dec.DecodeElement(&s, &t); err == nil {
						cellText.WriteString(s)
					}
					depth--
					continue
				}
			case inBody && inTable && inCell && local == "br":
				cellText.WriteString("\n")
			}
		case xml.EndElement:
			local := t.Name.Local
			switch {
			case local == "body":
				inBody = false
			case inBody && local == "p" && inParagraph && !inTable:
				text := paraText.String()
				if text != "" || paraStyle != "" {
					elements = append(elements, BodyElement{
						Index: idx,
						Type:  "paragraph",
						Paragraph: &Paragraph{
							Index: idx,
							Style: paraStyle,
							Text:  text,
						},
					})
					idx++
				}
				inParagraph = false
			case inBody && inTable && local == "tc":
				currentRow = append(currentRow, cellText.String())
				inCell = false
			case inBody && inTable && local == "tr":
				currentTableRows = append(currentTableRows, currentRow)
			case inBody && local == "tbl":
				elements = append(elements, BodyElement{
					Index: idx,
					Type:  "table",
					Table: &TableData{Rows: currentTableRows},
				})
				idx++
				inTable = false
			case inBody && inTable && local == "p":
				// end of paragraph inside table cell
			}
			depth--
		}
	}
	return elements, nil
}

// TableXML generates a <w:tbl> XML string from rows. First row is bold.
func TableXML(rows [][]string) string {
	fontRPr := `<w:rFonts w:ascii="` + defaultFontLatin + `" w:hAnsi="` + defaultFontLatin +
		`" w:eastAsia="` + defaultFontEastAsia + `" w:cs="` + defaultFontComplex + `"/>`
	var sb strings.Builder
	sb.WriteString("<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>")
	for i, row := range rows {
		sb.WriteString("<w:tr>")
		for _, cell := range row {
			sb.WriteString("<w:tc><w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/></w:tcPr><w:p>")
			if i == 0 {
				sb.WriteString("<w:r><w:rPr><w:b/>")
				sb.WriteString(fontRPr)
				sb.WriteString("</w:rPr><w:t xml:space=\"preserve\">")
			} else {
				sb.WriteString("<w:r><w:rPr>")
				sb.WriteString(fontRPr)
				sb.WriteString("</w:rPr><w:t xml:space=\"preserve\">")
			}
			sb.WriteString(xmlEscape(cell))
			sb.WriteString("</w:t></w:r></w:p></w:tc>")
		}
		sb.WriteString("</w:tr>")
	}
	sb.WriteString("</w:tbl>")
	return sb.String()
}

// ParagraphXML generates a <w:p> XML string from a Paragraph.
func ParagraphXML(p Paragraph) string {
	var sb strings.Builder
	sb.WriteString("<w:p>")
	if p.Style != "" {
		sb.WriteString(`<w:pPr><w:pStyle w:val="`)
		sb.WriteString(xmlEscape(p.Style))
		sb.WriteString(`"/></w:pPr>`)
	}
	// Explicit run-level font ensures consistent rendering even when
	// the consuming application ignores <w:docDefaults>.
	sb.WriteString(`<w:r><w:rPr><w:rFonts w:ascii="`)
	sb.WriteString(defaultFontLatin)
	sb.WriteString(`" w:hAnsi="`)
	sb.WriteString(defaultFontLatin)
	sb.WriteString(`" w:eastAsia="`)
	sb.WriteString(defaultFontEastAsia)
	sb.WriteString(`" w:cs="`)
	sb.WriteString(defaultFontComplex)
	sb.WriteString(`"/></w:rPr><w:t xml:space="preserve">`)
	sb.WriteString(xmlEscape(p.Text))
	sb.WriteString("</w:t></w:r></w:p>")
	return sb.String()
}

// bodyElementXML serializes one BodyElement back to XML.
func bodyElementXML(e BodyElement) string {
	if e.Type == "table" && e.Table != nil {
		return TableXML(e.Table.Rows)
	}
	if e.Paragraph != nil {
		return ParagraphXML(*e.Paragraph)
	}
	return ""
}

// serializeBodyElements converts a list of BodyElements to the XML content
// that should go inside <w:body>...</w:body>.
func serializeBodyElements(elements []BodyElement) string {
	var sb strings.Builder
	for _, e := range elements {
		sb.WriteString("\n")
		sb.WriteString(bodyElementXML(e))
	}
	sb.WriteString("\n")
	return sb.String()
}

// rewriteBody reads document.xml, replaces its body with the given elements, and
// writes the result to outPath.
func rewriteBody(path, outPath string, elements []BodyElement) error {
	zc, err := common.OpenReader(path)
	if err != nil {
		return err
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		return err
	}

	bodyOpen := "<w:body>"
	bodyClose := "</w:body>"
	startIdx := bytes.Index(doc, []byte(bodyOpen))
	endIdx := bytes.LastIndex(doc, []byte(bodyClose))
	if startIdx < 0 || endIdx < 0 {
		return fmt.Errorf("invalid document.xml: missing <w:body>")
	}

	var buf bytes.Buffer
	buf.Write(doc[:startIdx+len(bodyOpen)])
	buf.WriteString(serializeBodyElements(elements))
	buf.Write(doc[endIdx:])

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		documentXMLPath: buf.Bytes(),
	})
}

// ---------------------------------------------------------------------------
// Structural operations: delete, insert, update-table-cell
// ---------------------------------------------------------------------------

// DeleteBodyElement removes the body element at the given 0-based index.
func DeleteBodyElement(path, outPath string, index int) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if index < 0 || index >= len(elements) {
		return fmt.Errorf("index %d out of range (document has %d body elements)", index, len(elements))
	}
	elements = append(elements[:index], elements[index+1:]...)
	// Re-index
	for i := range elements {
		elements[i].Index = i
		if elements[i].Paragraph != nil {
			elements[i].Paragraph.Index = i
		}
	}
	return rewriteBody(path, outPath, elements)
}

// InsertBodyElement inserts a new XML element at the given position.
// If before=true, it is inserted before elements[index]; otherwise after it.
// xmlContent is the raw XML string (e.g. a <w:p> or <w:tbl>).
func InsertBodyElement(path, outPath string, index int, before bool, xmlContent string) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if index < 0 || index > len(elements) {
		return fmt.Errorf("index %d out of range (document has %d body elements)", index, len(elements))
	}

	// Build a temporary BodyElement for the new content.
	// We'll serialize the whole body, so we inject raw XML instead.
	// Use a special approach: rebuild body with raw XML at the right position.
	zc, err := common.OpenReader(path)
	if err != nil {
		return err
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		return err
	}

	bodyOpen := "<w:body>"
	bodyClose := "</w:body>"
	startIdx := bytes.Index(doc, []byte(bodyOpen))
	endIdx := bytes.LastIndex(doc, []byte(bodyClose))
	if startIdx < 0 || endIdx < 0 {
		return fmt.Errorf("invalid document.xml: missing <w:body>")
	}

	// Collect each element's raw XML by walking the body XML
	rawElements := collectRawBodyElements(doc[startIdx+len(bodyOpen) : endIdx])

	insertPos := index
	if !before {
		insertPos = index + 1
	}
	if insertPos > len(rawElements) {
		insertPos = len(rawElements)
	}

	// Rebuild body
	var buf bytes.Buffer
	buf.Write(doc[:startIdx+len(bodyOpen)])
	for i, raw := range rawElements {
		if i == insertPos {
			buf.WriteString("\n")
			buf.WriteString(xmlContent)
		}
		buf.WriteString("\n")
		buf.Write(raw)
	}
	if insertPos >= len(rawElements) {
		buf.WriteString("\n")
		buf.WriteString(xmlContent)
	}
	buf.WriteString("\n")
	buf.Write(doc[endIdx:])

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		documentXMLPath: buf.Bytes(),
	})
}

// collectRawBodyElements splits the content between <w:body> and </w:body> into
// top-level elements by tracking <w:p>…</w:p> and <w:tbl>…</w:tbl> boundaries.
func collectRawBodyElements(bodyContent []byte) [][]byte {
	var elements [][]byte
	depth := 0
	var start int
	inElement := false
	s := string(bodyContent)

	for i := 0; i < len(s); {
		// Look for <w:p or <w:tbl (start of a top-level element)
		if !inElement {
			// Skip whitespace
			for i < len(s) && (s[i] == ' ' || s[i] == '\n' || s[i] == '\r' || s[i] == '\t') {
				i++
			}
			if i >= len(s) {
				break
			}
			// Check for opening tags
			if matchesTag(s, i, "w:p") || matchesTag(s, i, "w:tbl") {
				start = i
				inElement = true
				depth = 1
				// Skip past the opening tag
				for i < len(s) && s[i] != '>' {
					i++
				}
				if i < len(s) && s[i-1] == '/' {
					depth = 0
				}
				i++
				if depth == 0 {
					elements = append(elements, []byte(s[start:i]))
					inElement = false
				}
				continue
			}
			i++
			continue
		}

		// Inside an element: track depth
		if s[i] == '<' {
			j := i + 1
			var curTag string
			if j < len(s) && s[j] == '/' {
				// Closing tag — extract tag name
				j++
				tagStart := j
				for j < len(s) && s[j] != ' ' && s[j] != '>' && s[j] != '/' {
					j++
				}
				curTag = s[tagStart:j]
				if isBodyChild(curTag) {
					depth--
					if depth == 0 {
						for j < len(s) && s[j] != '>' {
							j++
						}
						j++
						elements = append(elements, []byte(s[start:j]))
						inElement = false
						i = j
						continue
					}
				}
			} else if j < len(s) && s[j] != '!' && s[j] != '?' {
				// Opening tag — check if it's one of our tracked types
				tagStart := j
				for j < len(s) && s[j] != ' ' && s[j] != '>' && s[j] != '/' && s[j] != '\n' {
					j++
				}
				curTag = s[tagStart:j]
				if isBodyChild(curTag) {
					depth++
				}
			}
			// Skip to '>'
			for i < len(s) && s[i] != '>' {
				i++
			}
			if i < len(s) && s[i-1] == '/' && inElement {
				// Self-closing tag — only adjust depth for body children
				if isBodyChild(curTag) {
					depth--
					if depth == 0 {
						i++
						elements = append(elements, []byte(s[start:i]))
						inElement = false
						continue
					}
				}
			}
		}
		i++
	}
	return elements
}

// matchesTag checks if s at position pos starts with "<tagName" (opening) or
// matches a self-closing or opening tag pattern.
func matchesTag(s string, pos int, tag string) bool {
	if pos+1+len(tag) > len(s) {
		return false
	}
	if s[pos] != '<' {
		return false
	}
	rest := s[pos+1:]
	if !strings.HasPrefix(rest, tag) {
		return false
	}
	after := rest[len(tag):]
	if len(after) == 0 {
		return true
	}
	c := after[0]
	return c == '>' || c == ' ' || c == '/' || c == '\n' || c == '\r' || c == '\t'
}

// isBodyChild reports whether a tag name is a direct child of <w:body>.
func isBodyChild(tag string) bool {
	return tag == "w:p" || tag == "w:tbl" || tag == "w:sdt" || tag == "w:sectPr"
}
