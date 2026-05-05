package word

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

type RunStyle struct {
	Bold       bool   `json:"bold,omitempty"`
	Italic     bool   `json:"italic,omitempty"`
	Underline  bool   `json:"underline,omitempty"`
	Strike     bool   `json:"strike,omitempty"`
	FontFamily string `json:"fontFamily,omitempty"`
	FontSize   int    `json:"fontSize,omitempty"` // in points (converted to half-points internally)
	Color      string `json:"color,omitempty"`    // hex without #
}

// IsZero reports whether no styling was requested.
func (s RunStyle) IsZero() bool {
	return !s.Bold && !s.Italic && !s.Underline && !s.Strike &&
		s.FontFamily == "" && s.FontSize == 0 && s.Color == ""
}

// ParaStyle describes paragraph-level formatting.
type ParaStyle struct {
	Align       string `json:"align,omitempty"`       // left, center, right, both
	SpaceBefore int    `json:"spaceBefore,omitempty"` // twips
	SpaceAfter  int    `json:"spaceAfter,omitempty"`  // twips
	LineSpacing int    `json:"lineSpacing,omitempty"` // twips (240=single, 360=1.5, 480=double)
	IndentLeft  int    `json:"indentLeft,omitempty"`  // twips
	IndentRight int    `json:"indentRight,omitempty"` // twips
	FirstLine   int    `json:"firstLine,omitempty"`   // twips
}

// IsZero reports whether no styling was requested.
func (s ParaStyle) IsZero() bool {
	return s.Align == "" && s.SpaceBefore == 0 && s.SpaceAfter == 0 &&
		s.LineSpacing == 0 && s.IndentLeft == 0 && s.IndentRight == 0 && s.FirstLine == 0
}

// CellStyle describes table cell formatting.
type CellStyle struct {
	BgColor string `json:"bgColor,omitempty"` // hex without #
	Borders string `json:"borders,omitempty"` // "none", "single", "thick"
	VAlign  string `json:"vAlign,omitempty"`  // top, center, bottom
}

// IsZero reports whether no styling was requested.
func (s CellStyle) IsZero() bool {
	return s.BgColor == "" && s.Borders == "" && s.VAlign == ""
}

// buildRunPropsXML generates <w:rPr>...</w:rPr> from a RunStyle.
func buildRunPropsXML(s RunStyle) string {
	if s.IsZero() {
		return ""
	}
	var sb strings.Builder
	sb.WriteString("<w:rPr>")
	if s.Bold {
		sb.WriteString("<w:b/>")
	}
	if s.Italic {
		sb.WriteString("<w:i/>")
	}
	if s.Underline {
		sb.WriteString(`<w:u w:val="single"/>`)
	}
	if s.Strike {
		sb.WriteString("<w:strike/>")
	}
	if s.FontFamily != "" {
		// Set all four font slots so the same font applies regardless of
		// character class (ASCII, high-ANSI, East Asian, complex scripts).
		f := xmlEscape(s.FontFamily)
		sb.WriteString(`<w:rFonts w:ascii="`)
		sb.WriteString(f)
		sb.WriteString(`" w:hAnsi="`)
		sb.WriteString(f)
		sb.WriteString(`" w:eastAsia="`)
		sb.WriteString(f)
		sb.WriteString(`" w:cs="`)
		sb.WriteString(f)
		sb.WriteString(`"/>`)
	}
	if s.FontSize > 0 {
		// OOXML uses half-points: 14pt = 28
		hp := s.FontSize * 2
		sb.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, hp))
	}
	if s.Color != "" {
		sb.WriteString(`<w:color w:val="`)
		sb.WriteString(xmlEscape(strings.TrimPrefix(s.Color, "#")))
		sb.WriteString(`"/>`)
	}
	sb.WriteString("</w:rPr>")
	return sb.String()
}

// buildParaPropsXML generates <w:pPr>...</w:pPr> from a ParaStyle.
func buildParaPropsXML(s ParaStyle) string {
	if s.IsZero() {
		return ""
	}
	var sb strings.Builder
	sb.WriteString("<w:pPr>")
	if s.Align != "" {
		sb.WriteString(fmt.Sprintf(`<w:jc w:val="%s"/>`, xmlEscape(s.Align)))
	}
	if s.SpaceBefore > 0 || s.SpaceAfter > 0 || s.LineSpacing > 0 {
		sb.WriteString("<w:spacing")
		if s.SpaceBefore > 0 {
			sb.WriteString(fmt.Sprintf(` w:before="%d"`, s.SpaceBefore))
		}
		if s.SpaceAfter > 0 {
			sb.WriteString(fmt.Sprintf(` w:after="%d"`, s.SpaceAfter))
		}
		if s.LineSpacing > 0 {
			sb.WriteString(fmt.Sprintf(` w:line="%d"`, s.LineSpacing))
		}
		sb.WriteString("/>")
	}
	if s.IndentLeft > 0 || s.IndentRight > 0 || s.FirstLine > 0 {
		sb.WriteString("<w:ind")
		if s.IndentLeft > 0 {
			sb.WriteString(fmt.Sprintf(` w:left="%d"`, s.IndentLeft))
		}
		if s.IndentRight > 0 {
			sb.WriteString(fmt.Sprintf(` w:right="%d"`, s.IndentRight))
		}
		if s.FirstLine > 0 {
			sb.WriteString(fmt.Sprintf(` w:firstLine="%d"`, s.FirstLine))
		}
		sb.WriteString("/>")
	}
	sb.WriteString("</w:pPr>")
	return sb.String()
}

// buildCellPropsXML generates <w:tcPr>...</w:tcPr> from a CellStyle.
func buildCellPropsXML(s CellStyle) string {
	if s.IsZero() {
		return ""
	}
	var sb strings.Builder
	sb.WriteString("<w:tcPr>")
	if s.BgColor != "" {
		sb.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="`)
		sb.WriteString(xmlEscape(strings.TrimPrefix(s.BgColor, "#")))
		sb.WriteString(`"/>`)
	}
	if s.Borders != "" {
		switch s.Borders {
		case "none":
			sb.WriteString(`<w:tcBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>`)
		case "thick":
			sb.WriteString(`<w:tcBorders><w:top w:val="thick" w:sz="12"/><w:left w:val="thick" w:sz="12"/><w:bottom w:val="thick" w:sz="12"/><w:right w:val="thick" w:sz="12"/></w:tcBorders>`)
		default: // "single"
			sb.WriteString(`<w:tcBorders><w:top w:val="single"/><w:left w:val="single"/><w:bottom w:val="single"/><w:right w:val="single"/></w:tcBorders>`)
		}
	}
	if s.VAlign != "" {
		sb.WriteString(fmt.Sprintf(`<w:vAlign w:val="%s"/>`, xmlEscape(s.VAlign)))
	}
	sb.WriteString("</w:tcPr>")
	return sb.String()
}

// replaceOrInsertBlock replaces an existing XML block (from openTag to closeTag)
// with newContent. If the block doesn't exist, inserts newContent after the parent
// opening tag. Returns the modified byte slice.
func replaceOrInsertBlock(raw []byte, openTag, closeTag, newContent string, parentTag string) []byte {
	s := string(raw)
	openIdx := strings.Index(s, openTag)
	if openIdx >= 0 {
		// Find the matching close tag
		closeIdx := strings.Index(s[openIdx:], closeTag)
		if closeIdx >= 0 {
			closeIdx += openIdx + len(closeTag)
			// Replace the entire block
			var buf bytes.Buffer
			buf.WriteString(s[:openIdx])
			buf.WriteString(newContent)
			buf.WriteString(s[closeIdx:])
			return buf.Bytes()
		}
	}
	// Not found — insert after the parent opening tag
	parentIdx := strings.Index(s, parentTag)
	if parentIdx >= 0 {
		insertAt := parentIdx + len(parentTag)
		var buf bytes.Buffer
		buf.WriteString(s[:insertAt])
		buf.WriteString(newContent)
		buf.WriteString(s[insertAt:])
		return buf.Bytes()
	}
	return raw
}

// injectRunProps replaces or inserts <w:rPr> in a <w:r> element.
func injectRunProps(raw []byte, props string) []byte {
	return replaceOrInsertBlock(raw, "<w:rPr>", "</w:rPr>", props, "<w:r>")
}

// injectParaProps replaces or inserts <w:pPr> in a <w:p> element.
func injectParaProps(raw []byte, props string) []byte {
	return replaceOrInsertBlock(raw, "<w:pPr>", "</w:pPr>", props, "<w:p>")
}

// injectCellProps replaces or inserts <w:tcPr> in a <w:tc> element.
func injectCellProps(raw []byte, props string) []byte {
	return replaceOrInsertBlock(raw, "<w:tcPr>", "</w:tcPr>", props, "<w:tc>")
}

// applyRunStyleToAllRuns applies RunStyle to every <w:r> in the given XML.
func applyRunStyleToAllRuns(raw []byte, style RunStyle) []byte {
	if style.IsZero() {
		return raw
	}
	props := buildRunPropsXML(style)
	s := string(raw)
	var result bytes.Buffer
	pos := 0
	for {
		idx := strings.Index(s[pos:], "<w:r>")
		if idx < 0 {
			result.WriteString(s[pos:])
			break
		}
		idx += pos
		// Find the end of this <w:r>…</w:r>
		endIdx := strings.Index(s[idx:], "</w:r>")
		if endIdx < 0 {
			result.WriteString(s[pos:])
			break
		}
		endIdx += idx + len("</w:r>")
		run := []byte(s[idx:endIdx])
		styled := injectRunProps(run, props)
		result.WriteString(s[pos:idx])
		result.Write(styled)
		pos = endIdx
	}
	return result.Bytes()
}

// ---------------------------------------------------------------------------
// High-level style functions
// ---------------------------------------------------------------------------

// readRawBody reads document.xml, extracts the body content, and splits it
// into raw element XML bytes. Returns the full document, the body opening tag
// string (including the newline), and the list of raw elements.
func readRawBody(path string) (fullDoc []byte, bodyPrefix string, rawElements [][]byte, err error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return nil, "", nil, err
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		return nil, "", nil, err
	}

	bodyOpen := "<w:body>"
	bodyClose := "</w:body>"
	startIdx := bytes.Index(doc, []byte(bodyOpen))
	endIdx := bytes.LastIndex(doc, []byte(bodyClose))
	if startIdx < 0 || endIdx < 0 {
		return nil, "", nil, fmt.Errorf("invalid document.xml: missing <w:body>")
	}

	bodyContent := doc[startIdx+len(bodyOpen) : endIdx]
	rawElements = collectRawBodyElements(bodyContent)
	return doc, string(doc[:startIdx+len(bodyOpen)]), rawElements, nil
}

// writeRawBody reassembles the full document from raw elements and writes to
// outPath (or back to path if outPath is empty).
func writeRawBody(path, outPath string, fullDoc []byte, bodyPrefix string, rawElements [][]byte) error {
	var buf bytes.Buffer
	buf.WriteString(bodyPrefix)
	for _, raw := range rawElements {
		buf.WriteString("\n")
		buf.Write(raw)
	}
	buf.WriteString("\n")
	// Find the closing body tag position in the original doc
	bodyCloseTag := "</w:body>"
	closeIdx := bytes.LastIndex(fullDoc, []byte(bodyCloseTag))
	if closeIdx < 0 {
		return fmt.Errorf("invalid document.xml: missing </w:body>")
	}
	buf.Write(fullDoc[closeIdx:])

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		documentXMLPath: buf.Bytes(),
	})
}

// StyleParagraph applies styling to the paragraph at the given body-element
// index. runStyle is applied to every run inside the paragraph; paraStyle is
// applied to the paragraph itself.
func StyleParagraph(path, outPath string, index int, runStyle RunStyle, paraStyle ParaStyle) error {
	fullDoc, prefix, rawElements, err := readRawBody(path)
	if err != nil {
		return err
	}
	if index < 0 || index >= len(rawElements) {
		return fmt.Errorf("element index %d out of range (document has %d body elements)", index, len(rawElements))
	}

	elem := rawElements[index]
	// Verify it's a paragraph
	if !bytes.HasPrefix(bytes.TrimSpace(elem), []byte("<w:p")) &&
		!bytes.HasPrefix(bytes.TrimSpace(elem), []byte("<w:p>")) {
		return fmt.Errorf("body element %d is not a paragraph", index)
	}

	// Apply paragraph-level style
	if !paraStyle.IsZero() {
		pprXML := buildParaPropsXML(paraStyle)
		elem = injectParaProps(elem, pprXML)
	}
	// Apply run-level style to all runs
	if !runStyle.IsZero() {
		elem = applyRunStyleToAllRuns(elem, runStyle)
	}

	rawElements[index] = elem
	return writeRawBody(path, outPath, fullDoc, prefix, rawElements)
}

// StyleTable applies styling to a range of cells in the table at the given
// body-element index. cellStyle is applied to each cell; runStyle is applied
// to every run inside each cell. Use -1 for startRow/endRow/startCol/endCol
// to mean "all".
func StyleTable(path, outPath string, tableIndex int, startRow, endRow, startCol, endCol int, cellStyle CellStyle, runStyle RunStyle) error {
	fullDoc, prefix, rawElements, err := readRawBody(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(rawElements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(rawElements))
	}

	tblElem := rawElements[tableIndex]
	s := string(tblElem)
	// Must be a table
	if !bytes.HasPrefix(bytes.TrimSpace(tblElem), []byte("<w:tbl")) {
		return fmt.Errorf("body element %d is not a table", tableIndex)
	}

	// Find <w:tbl>...</w:tbl> — may be preceded by <w:p> (paragraph) wrapping,
	// but collectRawBodyElements should give us the raw table element.
	tblOpen := "<w:tbl>"
	tblClose := "</w:tbl>"
	tblStart := strings.Index(s, tblOpen)
	tblEnd := strings.LastIndex(s, tblClose)
	if tblStart < 0 || tblEnd < 0 {
		return fmt.Errorf("body element %d: cannot find <w:tbl> boundaries", tableIndex)
	}
	tblEnd += len(tblClose)

	// Split table into rows: find all <w:tr>...</w:tr>
	rows := splitXMLBlocks(s[tblStart:tblEnd], "<w:tr>", "</w:tr>")
	if len(rows) == 0 {
		return fmt.Errorf("table at element %d has no rows", tableIndex)
	}

	// Normalize range
	if startRow < 0 {
		startRow = 0
	}
	if endRow < 0 || endRow >= len(rows) {
		endRow = len(rows) - 1
	}

	for r := startRow; r <= endRow; r++ {
		// Split row into cells: find all <w:tc>...</w:tc>
		cells := splitXMLBlocks(string(rows[r]), "<w:tc>", "</w:tc>")
		if len(cells) == 0 {
			continue
		}
		sc, ec := startCol, endCol
		if sc < 0 {
			sc = 0
		}
		if ec < 0 || ec >= len(cells) {
			ec = len(cells) - 1
		}
		for c := sc; c <= ec; c++ {
			cell := cells[c]
			if !cellStyle.IsZero() {
				cell = injectCellProps(cell, buildCellPropsXML(cellStyle))
			}
			if !runStyle.IsZero() {
				cell = applyRunStyleToAllRuns(cell, runStyle)
			}
			cells[c] = cell
		}
		// Reassemble row
		rows[r] = reassembleBlocks(string(rows[r]), "<w:tc>", "</w:tc>", cells)
	}

	// Reassemble table
	newTbl := reassembleBlocks(s[tblStart:tblEnd], "<w:tr>", "</w:tr>", rows)

	// Put the styled table back into the element
	var newElem bytes.Buffer
	newElem.WriteString(s[:tblStart])
	newElem.Write(newTbl)
	newElem.WriteString(s[tblEnd:])
	rawElements[tableIndex] = newElem.Bytes()

	return writeRawBody(path, outPath, fullDoc, prefix, rawElements)
}

// splitXMLBlocks splits a string into non-overlapping blocks delimited by
// openTag and closeTag. Each returned slice element is the raw bytes of one
// block including its open and close tags.
func splitXMLBlocks(s string, openTag, closeTag string) [][]byte {
	var blocks [][]byte
	pos := 0
	for {
		start := strings.Index(s[pos:], openTag)
		if start < 0 {
			break
		}
		start += pos
		end := strings.Index(s[start:], closeTag)
		if end < 0 {
			break
		}
		end += start + len(closeTag)
		blocks = append(blocks, []byte(s[start:end]))
		pos = end
	}
	return blocks
}

// reassembleBlocks reconstructs a string by replacing each occurrence of
// openTag…closeTag in order with the corresponding element from blocks.
func reassembleBlocks(s string, openTag, closeTag string, blocks [][]byte) []byte {
	var buf bytes.Buffer
	pos := 0
	idx := 0
	for idx < len(blocks) {
		start := strings.Index(s[pos:], openTag)
		if start < 0 {
			break
		}
		start += pos
		end := strings.Index(s[start:], closeTag)
		if end < 0 {
			break
		}
		end += start + len(closeTag)
		buf.WriteString(s[pos:start])
		buf.Write(blocks[idx])
		pos = end
		idx++
	}
	buf.WriteString(s[pos:])
	return buf.Bytes()
}

// UpdateTableCell modifies a single cell in a table.
// tableIndex is the body-element index of the table, row/col are 0-based.
func UpdateTableCell(path, outPath string, tableIndex, row, col int, value string) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	if row < 0 || row >= len(elem.Table.Rows) {
		return fmt.Errorf("row %d out of range (table has %d rows)", row, len(elem.Table.Rows))
	}
	if col < 0 || col >= len(elem.Table.Rows[row]) {
		return fmt.Errorf("col %d out of range (row %d has %d columns)", col, row, len(elem.Table.Rows[row]))
	}
	elem.Table.Rows[row][col] = value
	return rewriteBody(path, outPath, elements)
}

// ---------------------------------------------------------------------------
// Merge: combine multiple .docx files into one
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Search: keyword search across all body elements (paragraphs + table cells)
// ---------------------------------------------------------------------------

// SearchResult is one hit from a body-element search.
type SearchResult struct {
	ElementIndex int    `json:"index"`
	ElementType  string `json:"type"` // "paragraph" or "table"
	// For paragraphs:
	Style string `json:"style,omitempty"`
	Text  string `json:"text,omitempty"`
	// For tables — which row/col matched:
	Row  int        `json:"row,omitempty"`
	Col  int        `json:"col,omitempty"`
	Cell string     `json:"cell,omitempty"`
	Rows [][]string `json:"rows,omitempty"` // full table included when a table matches
}

// SearchBodyElements finds all body elements whose text contains keyword (case-insensitive).
// For paragraphs, the entire paragraph text is returned. For tables, the whole table
// is returned along with which specific cells matched.
func SearchBodyElements(path, keyword string) ([]SearchResult, error) {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return nil, err
	}
	needle := strings.ToLower(keyword)
	var results []SearchResult
	for _, e := range elements {
		if e.Paragraph != nil {
			if strings.Contains(strings.ToLower(e.Paragraph.Text), needle) {
				results = append(results, SearchResult{
					ElementIndex: e.Index,
					ElementType:  "paragraph",
					Style:        e.Paragraph.Style,
					Text:         e.Paragraph.Text,
				})
			}
		}
		if e.Table != nil {
			var matches []SearchResult
			for ri, row := range e.Table.Rows {
				for ci, cell := range row {
					if strings.Contains(strings.ToLower(cell), needle) {
						matches = append(matches, SearchResult{
							ElementIndex: e.Index,
							ElementType:  "table",
							Row:          ri,
							Col:          ci,
							Cell:         cell,
						})
					}
				}
			}
			if len(matches) > 0 {
				// Attach the full table rows to the first match; subsequent matches
				// reference the same element.
				matches[0].Rows = e.Table.Rows
				results = append(results, matches...)
			}
		}
	}
	return results, nil
}

// ---------------------------------------------------------------------------
// Table row/col operations (modify table in-place via body-element rewrite)
// ---------------------------------------------------------------------------

// AddTableRows inserts new rows into a table at the given body-element index.
// rows is the list of rows to insert; position is the 0-based row index to insert
// before (-1 = append at end).
func AddTableRows(path, outPath string, tableIndex, position int, rows [][]string) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	if position < 0 || position > len(elem.Table.Rows) {
		position = len(elem.Table.Rows) // append
	}
	// Normalize column count: ensure new rows match existing column count
	existingCols := 0
	if len(elem.Table.Rows) > 0 {
		existingCols = len(elem.Table.Rows[0])
	}
	for i, row := range rows {
		if existingCols > 0 && len(row) < existingCols {
			padded := make([]string, existingCols)
			copy(padded, row)
			rows[i] = padded
		}
	}
	// Insert
	newRows := make([][]string, 0, len(elem.Table.Rows)+len(rows))
	newRows = append(newRows, elem.Table.Rows[:position]...)
	newRows = append(newRows, rows...)
	newRows = append(newRows, elem.Table.Rows[position:]...)
	elem.Table.Rows = newRows
	return rewriteBody(path, outPath, elements)
}

// DeleteTableRows removes rows from a table. startRow and endRow are 0-based
// inclusive indices. To delete a single row, set startRow == endRow.
func DeleteTableRows(path, outPath string, tableIndex, startRow, endRow int) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	if startRow < 0 || startRow >= len(elem.Table.Rows) {
		return fmt.Errorf("start-row %d out of range (table has %d rows)", startRow, len(elem.Table.Rows))
	}
	if endRow < startRow || endRow >= len(elem.Table.Rows) {
		return fmt.Errorf("end-row %d out of range (table has %d rows)", endRow, len(elem.Table.Rows))
	}
	// Remove rows [startRow, endRow]
	elem.Table.Rows = append(elem.Table.Rows[:startRow], elem.Table.Rows[endRow+1:]...)
	return rewriteBody(path, outPath, elements)
}

// AddTableCols inserts new columns into a table at the given column position.
// defaultValues is a list of default values for the new column cells (one per row);
// if shorter than the number of rows, remaining cells are empty.
// position is the 0-based column index to insert before (-1 = append at end).
func AddTableCols(path, outPath string, tableIndex, position int, defaultValues []string) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	numRows := len(elem.Table.Rows)
	if numRows == 0 {
		return fmt.Errorf("table has no rows")
	}
	cols := len(elem.Table.Rows[0])
	if position < 0 || position > cols {
		position = cols // append
	}
	for i := range elem.Table.Rows {
		val := ""
		if i < len(defaultValues) {
			val = defaultValues[i]
		}
		newRow := make([]string, 0, cols+1)
		newRow = append(newRow, elem.Table.Rows[i][:position]...)
		newRow = append(newRow, val)
		newRow = append(newRow, elem.Table.Rows[i][position:]...)
		elem.Table.Rows[i] = newRow
	}
	return rewriteBody(path, outPath, elements)
}

// DeleteTableCols removes columns from a table. startCol and endCol are 0-based
// inclusive indices. To delete a single column, set startCol == endCol.
func DeleteTableCols(path, outPath string, tableIndex, startCol, endCol int) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	if len(elem.Table.Rows) == 0 {
		return fmt.Errorf("table has no rows")
	}
	cols := len(elem.Table.Rows[0])
	if startCol < 0 || startCol >= cols {
		return fmt.Errorf("start-col %d out of range (table has %d columns)", startCol, cols)
	}
	if endCol < startCol || endCol >= cols {
		return fmt.Errorf("end-col %d out of range (table has %d columns)", endCol, cols)
	}
	for i := range elem.Table.Rows {
		elem.Table.Rows[i] = append(elem.Table.Rows[i][:startCol], elem.Table.Rows[i][endCol+1:]...)
	}
	return rewriteBody(path, outPath, elements)
}

// CellUpdate describes a single cell update for UpdateTableCells.
type CellUpdate struct {
	Row   int    `json:"row"`
	Col   int    `json:"col"`
	Value string `json:"value"`
}

// UpdateTableCells performs batch updates to multiple cells in a single table.
// tableIndex is the body-element index of the table.
func UpdateTableCells(path, outPath string, tableIndex int, updates []CellUpdate) error {
	elements, err := ReadBodyElements(path)
	if err != nil {
		return err
	}
	if tableIndex < 0 || tableIndex >= len(elements) {
		return fmt.Errorf("table-index %d out of range (document has %d body elements)", tableIndex, len(elements))
	}
	elem := &elements[tableIndex]
	if elem.Type != "table" || elem.Table == nil {
		return fmt.Errorf("body element %d is not a table (type=%s)", tableIndex, elem.Type)
	}
	for _, u := range updates {
		if u.Row < 0 || u.Row >= len(elem.Table.Rows) {
			return fmt.Errorf("row %d out of range (table has %d rows)", u.Row, len(elem.Table.Rows))
		}
		if u.Col < 0 || u.Col >= len(elem.Table.Rows[u.Row]) {
			return fmt.Errorf("col %d out of range (row %d has %d columns)", u.Col, u.Row, len(elem.Table.Rows[u.Row]))
		}
		elem.Table.Rows[u.Row][u.Col] = u.Value
	}
	return rewriteBody(path, outPath, elements)
}

// Merge combines the body content of multiple .docx files into a single output file.
// The first file serves as the base (its styles, headers, footers are preserved).
// Media files from all sources are copied in.
func Merge(paths []string, outPath string) error {
	if len(paths) < 2 {
		return errors.New("merge requires at least 2 files")
	}

	// Read all body elements from all files
	var allElements []BodyElement
	for _, p := range paths {
		elements, err := ReadBodyElements(p)
		if err != nil {
			return fmt.Errorf("reading %s: %w", p, err)
		}
		allElements = append(allElements, elements...)
	}

	// Use the first file as the base
	base := paths[0]

	// Copy media from other files into the base
	// First, count existing media in base
	baseF, err := os.Open(base)
	if err != nil {
		return err
	}
	baseStat, err := baseF.Stat()
	if err != nil {
		_ = baseF.Close()
		return err
	}
	baseZr, err := zip.NewReader(baseF, baseStat.Size())
	if err != nil {
		_ = baseF.Close()
		return err
	}

	baseMediaCount := 0
	for _, f := range baseZr.File {
		if strings.HasPrefix(f.Name, "word/media/") {
			baseMediaCount++
		}
	}
	_ = baseF.Close()

	// Collect extra entries from other files (media, styles)
	extraEntries := map[string][]byte{}
	mediaOffset := baseMediaCount

	for i := 1; i < len(paths); i++ {
		f, err := os.Open(paths[i])
		if err != nil {
			return err
		}
		stat, err := f.Stat()
		if err != nil {
			_ = f.Close()
			return err
		}
		zr, err := zip.NewReader(f, stat.Size())
		if err != nil {
			_ = f.Close()
			return err
		}

		localMediaCount := 0
		for _, entry := range zr.File {
			if !strings.HasPrefix(entry.Name, "word/media/") {
				continue
			}
			localMediaCount++
			base := strings.TrimPrefix(entry.Name, "word/media/")
			if base == "" {
				continue
			}
			// Rename to avoid collision
			ext := filepath.Ext(base)
			newName := fmt.Sprintf("image%d%s", mediaOffset+localMediaCount, ext)
			rc, err := entry.Open()
			if err != nil {
				continue
			}
			data, err := io.ReadAll(rc)
			_ = rc.Close()
			if err != nil {
				continue
			}
			extraEntries["word/media/"+newName] = data
		}
		mediaOffset += localMediaCount
		_ = f.Close()
	}

	// Rewrite the base with merged body + extra media
	if err := rewriteBody(base, outPath, allElements); err != nil {
		return err
	}

	// Add extra media entries if any
	if len(extraEntries) > 0 {
		return common.RewriteEntries(outPath, outPath, extraEntries)
	}
	return nil
}
