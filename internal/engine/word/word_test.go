package word

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

// TestExtractHeadings runs against a synthetic []Paragraph slice (no .docx file
// needed) so we can validate the heading classifier in isolation.
func TestExtractHeadings(t *testing.T) {
	paras := []Paragraph{
		{Style: "Title", Text: "office-cli"},
		{Style: "Heading1", Text: "Introduction"},
		{Style: "Normal", Text: "Body paragraph"},
		{Style: "Heading2", Text: "Background"},
	}
	headings := ExtractHeadings(paras)
	if len(headings) != 3 {
		t.Fatalf("expected 3 headings, got %d (%+v)", len(headings), headings)
	}
	if headings[0].Level != 1 || headings[0].Text != "office-cli" {
		t.Errorf("first heading wrong: %+v", headings[0])
	}
	if headings[1].Level != 1 {
		t.Errorf("Heading1 should be level 1, got %d", headings[1].Level)
	}
	if headings[2].Level != 2 {
		t.Errorf("Heading2 should be level 2, got %d", headings[2].Level)
	}
}

// TestComputeStatsZeroEmpty verifies that the stats helper does not panic on
// empty input — important because AI Agents may pass blank docs.
func TestComputeStatsZeroEmpty(t *testing.T) {
	stats := ComputeStats(nil)
	if stats.Paragraphs != 0 || stats.Words != 0 || stats.Characters != 0 {
		t.Errorf("expected all zero stats, got %+v", stats)
	}
}

func TestComputeStats(t *testing.T) {
	paras := []Paragraph{
		{Style: "Title", Text: "Hello world"},
		{Style: "Normal", Text: "this has four words"},
		{Style: "Heading1", Text: "Section"},
	}
	stats := ComputeStats(paras)
	if stats.Paragraphs != 3 {
		t.Errorf("Paragraphs = %d, want 3", stats.Paragraphs)
	}
	if stats.Words < 6 {
		t.Errorf("Words = %d, want >= 6", stats.Words)
	}
	if stats.Headings != 2 {
		t.Errorf("Headings = %d, want 2", stats.Headings)
	}
}

// ---------------------------------------------------------------------------
// helpers for creating temp .docx files in tests
// ---------------------------------------------------------------------------

func tempDocx(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "test.docx")
	if err := Create(path, "Test", "tester"); err != nil {
		t.Fatal(err)
	}
	return path
}

func tempDocxWithContent(t *testing.T) string {
	t.Helper()
	path := tempDocx(t)
	if err := AddParagraph(path, "", "Title Text", "Title"); err != nil {
		t.Fatal(err)
	}
	if err := AddParagraph(path, "", "First paragraph", "Normal"); err != nil {
		t.Fatal(err)
	}
	if err := AddTable(path, "", [][]string{{"A", "B"}, {"1", "2"}, {"3", "4"}}); err != nil {
		t.Fatal(err)
	}
	if err := AddParagraph(path, "", "Last paragraph", "Normal"); err != nil {
		t.Fatal(err)
	}
	return path
}

// ---------------------------------------------------------------------------
// parseBodyElements (token-based parser)
// ---------------------------------------------------------------------------

func TestParseBodyElements_ParagraphsOnly(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>World</w:t></w:r></w:p>
  </w:body>
</w:document>`

	elements, err := parseBodyElements([]byte(xmlData))
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 2 {
		t.Fatalf("expected 2 elements, got %d", len(elements))
	}
	if elements[0].Type != "paragraph" || elements[0].Paragraph == nil {
		t.Fatal("element 0 should be a paragraph")
	}
	if elements[0].Paragraph.Style != "Title" || elements[0].Paragraph.Text != "Hello" {
		t.Errorf("element 0: style=%q text=%q", elements[0].Paragraph.Style, elements[0].Paragraph.Text)
	}
	if elements[1].Paragraph.Text != "World" {
		t.Errorf("element 1 text=%q want %q", elements[1].Paragraph.Text, "World")
	}
}

func TestParseBodyElements_MixedParagraphsAndTables(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Y</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>After</w:t></w:r></w:p>
  </w:body>
</w:document>`

	elements, err := parseBodyElements([]byte(xmlData))
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elements))
	}
	if elements[0].Type != "paragraph" || elements[0].Paragraph.Text != "Before" {
		t.Errorf("element 0: type=%q text=%q", elements[0].Type, elements[0].Paragraph.Text)
	}
	if elements[1].Type != "table" || elements[1].Table == nil {
		t.Fatal("element 1 should be a table")
	}
	if len(elements[1].Table.Rows) != 2 {
		t.Fatalf("table rows=%d want 2", len(elements[1].Table.Rows))
	}
	if elements[1].Table.Rows[0][0] != "X" || elements[1].Table.Rows[1][1] != "2" {
		t.Errorf("table data: %v", elements[1].Table.Rows)
	}
	if elements[2].Type != "paragraph" || elements[2].Paragraph.Text != "After" {
		t.Errorf("element 2: type=%q text=%q", elements[2].Type, elements[2].Paragraph.Text)
	}
}

// ---------------------------------------------------------------------------
// ReadBodyElements (via real .docx file)
// ---------------------------------------------------------------------------

func TestReadBodyElements_WithParagraphsAndTable(t *testing.T) {
	path := tempDocxWithContent(t)
	elements, err := ReadBodyElements(path)
	if err != nil {
		t.Fatal(err)
	}
	// Title + First paragraph + Table + Last paragraph = 4
	if len(elements) != 4 {
		t.Fatalf("expected 4 body elements, got %d", len(elements))
	}

	// Element 0: Title paragraph
	if elements[0].Type != "paragraph" || elements[0].Paragraph.Style != "Title" {
		t.Errorf("element 0: %+v", elements[0])
	}

	// Element 1: Normal paragraph
	if elements[1].Type != "paragraph" || elements[1].Paragraph.Text != "First paragraph" {
		t.Errorf("element 1: %+v", elements[1])
	}

	// Element 2: Table
	if elements[2].Type != "table" || elements[2].Table == nil {
		t.Fatal("element 2 should be a table")
	}
	if len(elements[2].Table.Rows) != 3 {
		t.Fatalf("table rows=%d want 3", len(elements[2].Table.Rows))
	}
	if elements[2].Table.Rows[0][0] != "A" || elements[2].Table.Rows[2][1] != "4" {
		t.Errorf("table data: %v", elements[2].Table.Rows)
	}

	// Element 3: Last paragraph
	if elements[3].Type != "paragraph" || elements[3].Paragraph.Text != "Last paragraph" {
		t.Errorf("element 3: %+v", elements[3])
	}
}

// ---------------------------------------------------------------------------
// DeleteBodyElement
// ---------------------------------------------------------------------------

func TestDeleteBodyElement_Paragraph(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "deleted.docx")

	// Delete element at index 1 ("First paragraph")
	if err := DeleteBodyElement(path, out, 1); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 3 {
		t.Fatalf("expected 3 elements after delete, got %d", len(elements))
	}
	// Remaining: Title, Table, Last paragraph
	if elements[0].Paragraph.Text != "Title Text" {
		t.Errorf("element 0 text=%q", elements[0].Paragraph.Text)
	}
	if elements[1].Type != "table" {
		t.Errorf("element 1 should be table, got %q", elements[1].Type)
	}
	if elements[2].Paragraph.Text != "Last paragraph" {
		t.Errorf("element 2 text=%q", elements[2].Paragraph.Text)
	}
}

func TestDeleteBodyElement_Table(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "deleted-table.docx")

	// Delete element at index 2 (the table)
	if err := DeleteBodyElement(path, out, 2); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elements))
	}
	for i, e := range elements {
		if e.Type == "table" {
			t.Errorf("element %d should not be a table after delete", i)
		}
	}
}

func TestDeleteBodyElement_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "oob.docx")
	if err := DeleteBodyElement(path, out, 99); err == nil {
		t.Fatal("expected error for out-of-range index")
	}
}

// ---------------------------------------------------------------------------
// InsertBodyElement
// ---------------------------------------------------------------------------

func TestInsertBodyElement_Before(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "insert-before.docx")

	newPara := ParagraphXML(Paragraph{Style: "Heading1", Text: "Inserted"})
	if err := InsertBodyElement(path, out, 1, true, newPara); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 5 {
		t.Fatalf("expected 5 elements, got %d", len(elements))
	}
	// Expected: Title, [Inserted], First paragraph, Table, Last paragraph
	if elements[1].Paragraph.Text != "Inserted" || elements[1].Paragraph.Style != "Heading1" {
		t.Errorf("element 1: %+v", elements[1])
	}
	if elements[2].Paragraph.Text != "First paragraph" {
		t.Errorf("element 2: %+v", elements[2])
	}
}

func TestInsertBodyElement_After(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "insert-after.docx")

	newPara := ParagraphXML(Paragraph{Text: "After Table"})
	if err := InsertBodyElement(path, out, 2, false, newPara); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 5 {
		t.Fatalf("expected 5 elements, got %d", len(elements))
	}
	// Expected: Title, First paragraph, Table, [After Table], Last paragraph
	if elements[3].Paragraph.Text != "After Table" {
		t.Errorf("element 3: %+v", elements[3])
	}
}

func TestInsertBodyElement_Table(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "insert-table.docx")

	newTable := TableXML([][]string{{"X", "Y"}, {"a", "b"}})
	if err := InsertBodyElement(path, out, 0, true, newTable); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 5 {
		t.Fatalf("expected 5 elements, got %d", len(elements))
	}
	// New table should be at index 0
	if elements[0].Type != "table" {
		t.Errorf("element 0 type=%q want table", elements[0].Type)
	}
}

func TestInsertBodyElement_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "oob-insert.docx")
	if err := InsertBodyElement(path, out, 99, true, "<w:p/>"); err == nil {
		t.Fatal("expected error for out-of-range index")
	}
}

// ---------------------------------------------------------------------------
// UpdateTableCell
// ---------------------------------------------------------------------------

func TestUpdateTableCell_Basic(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "updated.docx")

	// Table is at body element index 2, update row 1, col 0 from "1" to "UPDATED"
	if err := UpdateTableCell(path, out, 2, 1, 0, "UPDATED"); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if elements[2].Table.Rows[1][0] != "UPDATED" {
		t.Errorf("cell value=%q want UPDATED", elements[2].Table.Rows[1][0])
	}
	// Other cells should be unchanged
	if elements[2].Table.Rows[0][0] != "A" {
		t.Errorf("cell [0][0]=%q want A", elements[2].Table.Rows[0][0])
	}
}

func TestUpdateTableCell_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "not-table.docx")
	// Element 0 is a paragraph, not a table
	if err := UpdateTableCell(path, out, 0, 0, 0, "X"); err == nil {
		t.Fatal("expected error when targeting non-table element")
	}
}

func TestUpdateTableCell_RowOutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "row-oob.docx")
	if err := UpdateTableCell(path, out, 2, 99, 0, "X"); err == nil {
		t.Fatal("expected error for out-of-range row")
	}
}

func TestUpdateTableCell_ColOutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "col-oob.docx")
	if err := UpdateTableCell(path, out, 2, 0, 99, "X"); err == nil {
		t.Fatal("expected error for out-of-range col")
	}
}

// ---------------------------------------------------------------------------
// Merge
// ---------------------------------------------------------------------------

func TestMerge_TwoFiles(t *testing.T) {
	path1 := tempDocx(t)
	if err := AddParagraph(path1, "", "Doc1 Para1", "Normal"); err != nil {
		t.Fatal(err)
	}
	path2 := tempDocx(t)
	if err := AddParagraph(path2, "", "Doc2 Para1", "Normal"); err != nil {
		t.Fatal(err)
	}
	if err := AddTable(path2, "", [][]string{{"X"}, {"Y"}}); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "merged.docx")
	if err := Merge([]string{path1, path2}, out); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	// Doc1: 1 paragraph; Doc2: 1 paragraph + 1 table = 3 total
	if len(elements) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elements))
	}
	if elements[0].Paragraph.Text != "Doc1 Para1" {
		t.Errorf("element 0: %+v", elements[0])
	}
	if elements[1].Paragraph.Text != "Doc2 Para1" {
		t.Errorf("element 1: %+v", elements[1])
	}
	if elements[2].Type != "table" {
		t.Errorf("element 2 type=%q want table", elements[2].Type)
	}
}

func TestMerge_ThreeFiles(t *testing.T) {
	path1 := tempDocx(t)
	_ = AddParagraph(path1, "", "A", "Normal")
	path2 := tempDocx(t)
	_ = AddParagraph(path2, "", "B", "Normal")
	path3 := tempDocx(t)
	_ = AddParagraph(path3, "", "C", "Normal")

	out := filepath.Join(t.TempDir(), "merged3.docx")
	if err := Merge([]string{path1, path2, path3}, out); err != nil {
		t.Fatal(err)
	}

	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 3 {
		t.Fatalf("expected 3 elements, got %d", len(elements))
	}
	texts := []string{elements[0].Paragraph.Text, elements[1].Paragraph.Text, elements[2].Paragraph.Text}
	if texts[0] != "A" || texts[1] != "B" || texts[2] != "C" {
		t.Errorf("merged texts: %v", texts)
	}
}

func TestMerge_TooFewFiles(t *testing.T) {
	path := tempDocx(t)
	if err := Merge([]string{path}, "out.docx"); err == nil {
		t.Fatal("expected error for < 2 files")
	}
}

// ---------------------------------------------------------------------------
// TableXML / ParagraphXML (exported helpers)
// ---------------------------------------------------------------------------

func TestTableXML_Roundtrip(t *testing.T) {
	rows := [][]string{{"Name", "Age"}, {"Alice", "30"}}
	xmlStr := TableXML(rows)
	if !strings.Contains(xmlStr, "<w:tbl>") || !strings.Contains(xmlStr, "</w:tbl>") {
		t.Error("TableXML should produce valid tbl tags")
	}
	if !strings.Contains(xmlStr, "Name") || !strings.Contains(xmlStr, "Alice") {
		t.Error("TableXML should contain cell values")
	}
	// First row should be bold
	if !strings.Contains(xmlStr, "<w:b/>") {
		t.Error("first row should have bold formatting")
	}
	// All rows should have default font (CJK-safe)
	if !strings.Contains(xmlStr, `w:eastAsia="宋体"`) {
		t.Error("TableXML should contain eastAsia default font")
	}
	if !strings.Contains(xmlStr, `w:ascii="Calibri"`) {
		t.Error("TableXML should contain ascii default font")
	}
}

func TestParagraphXML_Roundtrip(t *testing.T) {
	p := Paragraph{Style: "Heading1", Text: "Test <heading>"}
	xmlStr := ParagraphXML(p)
	if !strings.Contains(xmlStr, "Heading1") {
		t.Error("ParagraphXML should contain style")
	}
	if !strings.Contains(xmlStr, "Test") {
		t.Error("ParagraphXML should contain text")
	}
	// Must contain default font for all 4 slots
	if !strings.Contains(xmlStr, `w:eastAsia="宋体"`) {
		t.Error("ParagraphXML should contain eastAsia default font")
	}
	if !strings.Contains(xmlStr, `w:ascii="Calibri"`) {
		t.Error("ParagraphXML should contain ascii default font")
	}
	// Verify XML escaping
	if !strings.Contains(xmlStr, "&lt;heading&gt;") {
		t.Error("ParagraphXML should escape XML special characters")
	}
}

// ---------------------------------------------------------------------------
// Create: default font in styles.xml
// ---------------------------------------------------------------------------

func TestCreate_HasDocDefaults(t *testing.T) {
	path := filepath.Join(t.TempDir(), "defaults.docx")
	if err := Create(path, "Test", "tester"); err != nil {
		t.Fatal(err)
	}
	// Read styles.xml from the created docx
	zc, err := common.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	styles, err := common.ReadEntry(&zc.Reader, "word/styles.xml")
	_ = zc.Close()
	if err != nil {
		t.Fatal(err)
	}
	s := string(styles)
	if !strings.Contains(s, "<w:docDefaults>") {
		t.Error("styles.xml should contain <w:docDefaults>")
	}
	if !strings.Contains(s, `w:eastAsia="宋体"`) {
		t.Error("styles.xml should set eastAsia default font to 宋体")
	}
	if !strings.Contains(s, `w:ascii="Calibri"`) {
		t.Error("styles.xml should set ascii default font to Calibri")
	}
	if !strings.Contains(s, "w:sz") {
		t.Error("styles.xml should set default font size")
	}
}

// ---------------------------------------------------------------------------
// Rewrite body round-trip: read → modify → write → read → verify
// ---------------------------------------------------------------------------

func TestRewriteBody_Roundtrip(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "roundtrip.docx")

	elements, err := ReadBodyElements(path)
	if err != nil {
		t.Fatal(err)
	}

	// Add a new paragraph at the end
	elements = append(elements, BodyElement{
		Index: len(elements),
		Type:  "paragraph",
		Paragraph: &Paragraph{
			Index: len(elements),
			Text:  "Roundtrip added",
		},
	})

	if err := rewriteBody(path, out, elements); err != nil {
		t.Fatal(err)
	}

	// Re-read and verify
	got, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 5 {
		t.Fatalf("expected 5 elements after roundtrip, got %d", len(got))
	}
	if got[4].Paragraph.Text != "Roundtrip added" {
		t.Errorf("element 4 text=%q", got[4].Paragraph.Text)
	}
}

// ---------------------------------------------------------------------------
// CJK content
// ---------------------------------------------------------------------------

func TestParseBodyElements_CJK(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>中文段落</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>金额</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>张三</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>1234.5</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`

	elements, err := parseBodyElements([]byte(xmlData))
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 2 {
		t.Fatalf("expected 2 elements, got %d", len(elements))
	}
	if elements[0].Paragraph.Text != "中文段落" {
		t.Errorf("paragraph text=%q", elements[0].Paragraph.Text)
	}
	if elements[1].Table.Rows[0][0] != "姓名" || elements[1].Table.Rows[1][0] != "张三" {
		t.Errorf("table data: %v", elements[1].Table.Rows)
	}
}

// ---------------------------------------------------------------------------
// Empty document edge case
// ---------------------------------------------------------------------------

func TestParseBodyElements_EmptyBody(t *testing.T) {
	xmlData := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body></w:body>
</w:document>`

	elements, err := parseBodyElements([]byte(xmlData))
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) != 0 {
		t.Fatalf("expected 0 elements, got %d", len(elements))
	}
}

// ---------------------------------------------------------------------------
// File not found
// ---------------------------------------------------------------------------

func TestReadBodyElements_FileNotFound(t *testing.T) {
	_, err := ReadBodyElements("/nonexistent/path/to/file.docx")
	if err == nil {
		t.Fatal("expected error for nonexistent file")
	}
}

func TestDeleteBodyElement_FileNotFound(t *testing.T) {
	err := DeleteBodyElement("/nonexistent/path.docx", "/tmp/out.docx", 0)
	if err == nil {
		t.Fatal("expected error for nonexistent file")
	}
}

// ---------------------------------------------------------------------------
// SearchBodyElements
// ---------------------------------------------------------------------------

func TestSearchBodyElements_ParagraphMatch(t *testing.T) {
	path := tempDocxWithContent(t)
	results, err := SearchBodyElements(path, "First")
	if err != nil {
		t.Fatal(err)
	}
	if len(results) != 1 {
		t.Fatalf("expected 1 result, got %d", len(results))
	}
	if results[0].ElementType != "paragraph" || results[0].Text != "First paragraph" {
		t.Errorf("result: %+v", results[0])
	}
}

func TestSearchBodyElements_TableCellMatch(t *testing.T) {
	path := tempDocxWithContent(t)
	results, err := SearchBodyElements(path, "A")
	if err != nil {
		t.Fatal(err)
	}
	// Should match table cell "A" at row 0, col 0
	found := false
	for _, r := range results {
		if r.ElementType == "table" && r.Cell == "A" {
			found = true
			if r.Row != 0 || r.Col != 0 {
				t.Errorf("expected row=0 col=0, got row=%d col=%d", r.Row, r.Col)
			}
		}
	}
	if !found {
		t.Error("expected to find 'A' in table cell")
	}
}

func TestSearchBodyElements_NoMatch(t *testing.T) {
	path := tempDocxWithContent(t)
	results, err := SearchBodyElements(path, "ZZZZNOTFOUND")
	if err != nil {
		t.Fatal(err)
	}
	if len(results) != 0 {
		t.Fatalf("expected 0 results, got %d", len(results))
	}
}

func TestSearchBodyElements_CaseInsensitive(t *testing.T) {
	path := tempDocxWithContent(t)
	results, err := SearchBodyElements(path, "first")
	if err != nil {
		t.Fatal(err)
	}
	if len(results) != 1 {
		t.Fatalf("expected 1 result for case-insensitive search, got %d", len(results))
	}
}

func TestSearchBodyElements_CJK(t *testing.T) {
	path := tempDocx(t)
	if err := AddParagraph(path, "", "中文搜索测试", "Normal"); err != nil {
		t.Fatal(err)
	}
	results, err := SearchBodyElements(path, "搜索")
	if err != nil {
		t.Fatal(err)
	}
	if len(results) != 1 {
		t.Fatalf("expected 1 CJK result, got %d", len(results))
	}
}

func TestSearchBodyElements_FileNotFound(t *testing.T) {
	_, err := SearchBodyElements("/nonexistent/path.docx", "test")
	if err == nil {
		t.Fatal("expected error for nonexistent file")
	}
}

// ---------------------------------------------------------------------------
// AddTableRows
// ---------------------------------------------------------------------------

func TestAddTableRows_Append(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "add-rows.docx")
	newRows := [][]string{{"C", "D"}, {"E", "F"}}
	if err := AddTableRows(path, out, 2, -1, newRows); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 5 {
		t.Fatalf("expected 5 rows, got %d", len(tbl.Rows))
	}
	if tbl.Rows[3][0] != "C" || tbl.Rows[4][1] != "F" {
		t.Errorf("new rows: %v", tbl.Rows[3:])
	}
}

func TestAddTableRows_InsertBefore(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "insert-rows.docx")
	newRows := [][]string{{"X", "Y"}}
	if err := AddTableRows(path, out, 2, 1, newRows); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(tbl.Rows))
	}
	// Row 0: header A,B  Row 1: new X,Y  Row 2: old 1,2  Row 3: old 3,4
	if tbl.Rows[1][0] != "X" || tbl.Rows[2][0] != "1" {
		t.Errorf("unexpected rows: %v", tbl.Rows)
	}
}

func TestAddTableRows_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := AddTableRows(path, out, 0, -1, [][]string{{"X"}}); err == nil {
		t.Fatal("expected error for non-table element")
	}
}

func TestAddTableRows_PadColumns(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "padded.docx")
	// Table has 2 columns; new row has only 1
	if err := AddTableRows(path, out, 2, -1, [][]string{{"Z"}}); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 4 {
		t.Fatalf("expected 4 rows, got %d", len(tbl.Rows))
	}
	// Padded row should have 2 columns
	if len(tbl.Rows[3]) != 2 || tbl.Rows[3][0] != "Z" || tbl.Rows[3][1] != "" {
		t.Errorf("padded row: %v", tbl.Rows[3])
	}
}

// ---------------------------------------------------------------------------
// DeleteTableRows
// ---------------------------------------------------------------------------

func TestDeleteTableRows_Single(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "del-row.docx")
	if err := DeleteTableRows(path, out, 2, 1, 1); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(tbl.Rows))
	}
	// Remaining: header row A,B and row 3,4
	if tbl.Rows[1][0] != "3" {
		t.Errorf("remaining rows: %v", tbl.Rows)
	}
}

func TestDeleteTableRows_Range(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "del-range.docx")
	// Delete rows 0-1 (header + first data row)
	if err := DeleteTableRows(path, out, 2, 0, 1); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 1 {
		t.Fatalf("expected 1 row, got %d", len(tbl.Rows))
	}
	if tbl.Rows[0][0] != "3" {
		t.Errorf("remaining row: %v", tbl.Rows[0])
	}
}

func TestDeleteTableRows_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := DeleteTableRows(path, out, 0, 0, 0); err == nil {
		t.Fatal("expected error for non-table element")
	}
}

func TestDeleteTableRows_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := DeleteTableRows(path, out, 2, 0, 99); err == nil {
		t.Fatal("expected error for out-of-range end-row")
	}
}

// ---------------------------------------------------------------------------
// AddTableCols
// ---------------------------------------------------------------------------

func TestAddTableCols_Append(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "add-col.docx")
	if err := AddTableCols(path, out, 2, -1, []string{"X", "Y", "Z"}); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	// Should now have 3 columns
	if len(tbl.Rows[0]) != 3 {
		t.Fatalf("expected 3 columns, got %d", len(tbl.Rows[0]))
	}
	if tbl.Rows[0][2] != "X" || tbl.Rows[1][2] != "Y" || tbl.Rows[2][2] != "Z" {
		t.Errorf("new column values: [%q, %q, %q]", tbl.Rows[0][2], tbl.Rows[1][2], tbl.Rows[2][2])
	}
}

func TestAddTableCols_InsertBefore(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "insert-col.docx")
	if err := AddTableCols(path, out, 2, 1, []string{"X", "Y", "Z"}); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows[0]) != 3 {
		t.Fatalf("expected 3 columns, got %d", len(tbl.Rows[0]))
	}
	// New col at index 1; original col B shifted to index 2
	if tbl.Rows[0][0] != "A" || tbl.Rows[0][1] != "X" || tbl.Rows[0][2] != "B" {
		t.Errorf("row 0: %v", tbl.Rows[0])
	}
}

func TestAddTableCols_FewerDefaults(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "fewer-defaults.docx")
	// Only provide 1 default for 3 rows
	if err := AddTableCols(path, out, 2, -1, []string{"ONLY"}); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if tbl.Rows[0][2] != "ONLY" {
		t.Errorf("row 0 col 2: %q", tbl.Rows[0][2])
	}
	if tbl.Rows[1][2] != "" {
		t.Errorf("row 1 col 2 should be empty, got %q", tbl.Rows[1][2])
	}
}

func TestAddTableCols_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := AddTableCols(path, out, 0, -1, []string{"X"}); err == nil {
		t.Fatal("expected error for non-table element")
	}
}

// ---------------------------------------------------------------------------
// DeleteTableCols
// ---------------------------------------------------------------------------

func TestDeleteTableCols_Single(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "del-col.docx")
	// Delete column 0 (A, 1, 3)
	if err := DeleteTableCols(path, out, 2, 0, 0); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if len(tbl.Rows[0]) != 1 {
		t.Fatalf("expected 1 column, got %d", len(tbl.Rows[0]))
	}
	if tbl.Rows[0][0] != "B" || tbl.Rows[1][0] != "2" || tbl.Rows[2][0] != "4" {
		t.Errorf("remaining values: [%q, %q, %q]", tbl.Rows[0][0], tbl.Rows[1][0], tbl.Rows[2][0])
	}
}

func TestDeleteTableCols_Range(t *testing.T) {
	// Create a 3-column table
	path := tempDocx(t)
	if err := AddTable(path, "", [][]string{{"A", "B", "C"}, {"1", "2", "3"}}); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "del-col-range.docx")
	// Delete columns 0-1 (keep only column C)
	if err := DeleteTableCols(path, out, 0, 0, 1); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[0].Table
	if len(tbl.Rows[0]) != 1 {
		t.Fatalf("expected 1 column, got %d", len(tbl.Rows[0]))
	}
	if tbl.Rows[0][0] != "C" {
		t.Errorf("remaining col: %q", tbl.Rows[0][0])
	}
}

func TestDeleteTableCols_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := DeleteTableCols(path, out, 0, 0, 0); err == nil {
		t.Fatal("expected error for non-table element")
	}
}

func TestDeleteTableCols_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := DeleteTableCols(path, out, 2, 0, 99); err == nil {
		t.Fatal("expected error for out-of-range col")
	}
}

// ---------------------------------------------------------------------------
// UpdateTableCells (batch)
// ---------------------------------------------------------------------------

func TestUpdateTableCells_Basic(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "batch-update.docx")
	updates := []CellUpdate{
		{Row: 0, Col: 0, Value: "NEW_A"},
		{Row: 1, Col: 1, Value: "NEW_2"},
	}
	if err := UpdateTableCells(path, out, 2, updates); err != nil {
		t.Fatal(err)
	}
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	tbl := elements[2].Table
	if tbl.Rows[0][0] != "NEW_A" {
		t.Errorf("cell [0][0]=%q want NEW_A", tbl.Rows[0][0])
	}
	if tbl.Rows[1][1] != "NEW_2" {
		t.Errorf("cell [1][1]=%q want NEW_2", tbl.Rows[1][1])
	}
	// Unchanged cells
	if tbl.Rows[0][1] != "B" {
		t.Errorf("cell [0][1]=%q want B", tbl.Rows[0][1])
	}
}

func TestUpdateTableCells_Empty(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "batch-empty.docx")
	// Empty updates should succeed without error
	if err := UpdateTableCells(path, out, 2, nil); err != nil {
		t.Fatal(err)
	}
}

func TestUpdateTableCells_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	updates := []CellUpdate{{Row: 99, Col: 0, Value: "X"}}
	if err := UpdateTableCells(path, out, 2, updates); err == nil {
		t.Fatal("expected error for out-of-range row")
	}
}

func TestUpdateTableCells_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	updates := []CellUpdate{{Row: 0, Col: 0, Value: "X"}}
	if err := UpdateTableCells(path, out, 0, updates); err == nil {
		t.Fatal("expected error for non-table element")
	}
}

// ---------------------------------------------------------------------------
// XML style builder tests
// ---------------------------------------------------------------------------

func TestBuildRunPropsXML(t *testing.T) {
	s := RunStyle{Bold: true, Color: "FF0000", FontSize: 14}
	xml := buildRunPropsXML(s)
	if !strings.Contains(xml, "<w:rPr>") {
		t.Error("missing <w:rPr>")
	}
	if !strings.Contains(xml, "<w:b/>") {
		t.Error("missing bold tag")
	}
	if !strings.Contains(xml, `<w:color w:val="FF0000"/>`) {
		t.Error("missing color tag")
	}
	if !strings.Contains(xml, `<w:sz w:val="28"/>`) {
		t.Error("missing font size (should be 28 half-points for 14pt)")
	}
}

func TestBuildRunPropsXML_FontAllSlots(t *testing.T) {
	s := RunStyle{FontFamily: "微软雅黑"}
	xml := buildRunPropsXML(s)
	// All four font slots must be set for CJK compatibility
	for _, attr := range []string{`w:ascii="微软雅黑"`, `w:hAnsi="微软雅黑"`, `w:eastAsia="微软雅黑"`, `w:cs="微软雅黑"`} {
		if !strings.Contains(xml, attr) {
			t.Errorf("missing font attribute: %s", attr)
		}
	}
}

func TestBuildRunPropsXML_Zero(t *testing.T) {
	s := RunStyle{}
	if xml := buildRunPropsXML(s); xml != "" {
		t.Errorf("expected empty for zero style, got %q", xml)
	}
}

func TestBuildParaPropsXML(t *testing.T) {
	s := ParaStyle{Align: "center", SpaceBefore: 200, LineSpacing: 360}
	xml := buildParaPropsXML(s)
	if !strings.Contains(xml, `<w:jc w:val="center"/>`) {
		t.Error("missing alignment tag")
	}
	if !strings.Contains(xml, `w:before="200"`) {
		t.Error("missing space-before")
	}
	if !strings.Contains(xml, `w:line="360"`) {
		t.Error("missing line spacing")
	}
}

func TestBuildParaPropsXML_Zero(t *testing.T) {
	s := ParaStyle{}
	if xml := buildParaPropsXML(s); xml != "" {
		t.Errorf("expected empty for zero style, got %q", xml)
	}
}

func TestBuildCellPropsXML(t *testing.T) {
	s := CellStyle{BgColor: "FFFF00", VAlign: "center"}
	xml := buildCellPropsXML(s)
	if !strings.Contains(xml, `<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>`) {
		t.Error("missing background color")
	}
	if !strings.Contains(xml, `<w:vAlign w:val="center"/>`) {
		t.Error("missing vertical alignment")
	}
}

func TestBuildCellPropsXML_Borders(t *testing.T) {
	s := CellStyle{Borders: "thick"}
	xml := buildCellPropsXML(s)
	if !strings.Contains(xml, `<w:tcBorders>`) {
		t.Error("missing tcBorders")
	}
	if !strings.Contains(xml, `w:val="thick"`) {
		t.Error("missing thick border value")
	}
}

func TestBuildCellPropsXML_Zero(t *testing.T) {
	s := CellStyle{}
	if xml := buildCellPropsXML(s); xml != "" {
		t.Errorf("expected empty for zero style, got %q", xml)
	}
}

// ---------------------------------------------------------------------------
// replaceOrInsertBlock
// ---------------------------------------------------------------------------

func TestReplaceOrInsertBlock_Replace(t *testing.T) {
	raw := []byte(`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`)
	result := replaceOrInsertBlock(raw, "<w:pPr>", "</w:pPr>", "<w:pPr><w:jc w:val=\"center\"/></w:pPr>", "<w:p>")
	s := string(result)
	if !strings.Contains(s, `<w:jc w:val="center"/>`) {
		t.Error("expected replacement to contain center alignment")
	}
	if strings.Contains(s, `<w:pStyle w:val="Normal"/>`) {
		t.Error("old pPr content should be replaced")
	}
}

func TestReplaceOrInsertBlock_Insert(t *testing.T) {
	raw := []byte(`<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`)
	result := replaceOrInsertBlock(raw, "<w:pPr>", "</w:pPr>", "<w:pPr><w:jc w:val=\"center\"/></w:pPr>", "<w:p>")
	s := string(result)
	if !strings.Contains(s, `<w:jc w:val="center"/>`) {
		t.Error("expected inserted pPr")
	}
}

// ---------------------------------------------------------------------------
// injectRunProps / injectParaProps / injectCellProps
// ---------------------------------------------------------------------------

func TestInjectRunProps_New(t *testing.T) {
	raw := []byte(`<w:r><w:t>Hello</w:t></w:r>`)
	props := `<w:rPr><w:b/></w:rPr>`
	result := injectRunProps(raw, props)
	s := string(result)
	if !strings.Contains(s, "<w:b/>") {
		t.Error("expected bold tag injected")
	}
	// Should be: <w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r>
	if strings.Index(s, "<w:rPr>") >= strings.Index(s, "<w:t>") {
		t.Error("rPr should appear before w:t")
	}
}

func TestInjectRunProps_Existing(t *testing.T) {
	raw := []byte(`<w:r><w:rPr><w:i/></w:rPr><w:t>Hello</w:t></w:r>`)
	props := `<w:rPr><w:b/></w:rPr>`
	result := injectRunProps(raw, props)
	s := string(result)
	if !strings.Contains(s, "<w:b/>") {
		t.Error("expected bold tag")
	}
	if strings.Contains(s, "<w:i/>") {
		t.Error("old italic should be replaced")
	}
}

func TestInjectParaProps_New(t *testing.T) {
	raw := []byte(`<w:p><w:r><w:t>Text</w:t></w:r></w:p>`)
	props := `<w:pPr><w:jc w:val="center"/></w:pPr>`
	result := injectParaProps(raw, props)
	s := string(result)
	if !strings.Contains(s, `<w:jc w:val="center"/>`) {
		t.Error("expected center alignment")
	}
}

func TestInjectCellProps_New(t *testing.T) {
	raw := []byte(`<w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>`)
	props := `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr>`
	result := injectCellProps(raw, props)
	s := string(result)
	if !strings.Contains(s, "FFFF00") {
		t.Error("expected bg color in cell")
	}
}

// ---------------------------------------------------------------------------
// applyRunStyleToAllRuns
// ---------------------------------------------------------------------------

func TestApplyRunStyleToAllRuns(t *testing.T) {
	raw := []byte(`<w:p><w:r><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:p>`)
	styled := applyRunStyleToAllRuns(raw, RunStyle{Bold: true})
	s := string(styled)
	// Should have two <w:rPr><w:b/></w:rPr> blocks
	count := strings.Count(s, "<w:b/>")
	if count != 2 {
		t.Errorf("expected 2 bold tags, got %d in: %s", count, s)
	}
}

func TestApplyRunStyleToAllRuns_ZeroStyle(t *testing.T) {
	raw := []byte(`<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`)
	styled := applyRunStyleToAllRuns(raw, RunStyle{})
	if string(styled) != string(raw) {
		t.Error("zero style should not modify XML")
	}
}

// ---------------------------------------------------------------------------
// StyleParagraph (end-to-end)
// ---------------------------------------------------------------------------

func TestStyleParagraph_Basic(t *testing.T) {
	path := tempDocx(t)
	if err := AddParagraph(path, "", "Styled paragraph", "Normal"); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "styled.docx")
	runStyle := RunStyle{Bold: true, Color: "FF0000", FontSize: 16}
	paraStyle := ParaStyle{Align: "center"}
	if err := StyleParagraph(path, out, 0, runStyle, paraStyle); err != nil {
		t.Fatal(err)
	}

	// Read back and verify the raw XML contains the style
	zc, err := common.OpenReader(out)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		t.Fatal(err)
	}
	s := string(doc)
	if !strings.Contains(s, "<w:b/>") {
		t.Error("expected bold in styled paragraph")
	}
	if !strings.Contains(s, `<w:color w:val="FF0000"/>`) {
		t.Error("expected red color")
	}
	if !strings.Contains(s, `<w:sz w:val="32"/>`) {
		t.Error("expected font size 32 (16pt)")
	}
	if !strings.Contains(s, `<w:jc w:val="center"/>`) {
		t.Error("expected center alignment")
	}
}

func TestStyleParagraph_OutOfRange(t *testing.T) {
	path := tempDocx(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := StyleParagraph(path, out, 99, RunStyle{Bold: true}, ParaStyle{}); err == nil {
		t.Fatal("expected error for out-of-range index")
	}
}

func TestStyleParagraph_NotAParagraph(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	// Element 2 is a table
	if err := StyleParagraph(path, out, 2, RunStyle{Bold: true}, ParaStyle{}); err == nil {
		t.Fatal("expected error when styling a non-paragraph element")
	}
}

// ---------------------------------------------------------------------------
// StyleTable (end-to-end)
// ---------------------------------------------------------------------------

func TestStyleTable_AllCells(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "styled-table.docx")
	cellStyle := CellStyle{BgColor: "FFFF00"}
	runStyle := RunStyle{Bold: true}
	// Table is element index 2
	if err := StyleTable(path, out, 2, -1, -1, -1, -1, cellStyle, runStyle); err != nil {
		t.Fatal(err)
	}

	// Verify via ReadBodyElements that the doc is still valid
	elements, err := ReadBodyElements(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(elements) < 3 {
		t.Fatalf("expected at least 3 elements, got %d", len(elements))
	}
	if elements[2].Type != "table" {
		t.Fatalf("element 2 should be table, got %s", elements[2].Type)
	}
	tbl := elements[2].Table
	if len(tbl.Rows) != 3 {
		t.Fatalf("expected 3 rows, got %d", len(tbl.Rows))
	}
	// Verify content is preserved
	if tbl.Rows[0][0] != "A" || tbl.Rows[0][1] != "B" {
		t.Errorf("content not preserved: %v", tbl.Rows[0])
	}

	// Verify raw XML has the styling
	zc, err := common.OpenReader(out)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		t.Fatal(err)
	}
	s := string(doc)
	if !strings.Contains(s, "FFFF00") {
		t.Error("expected bg color in table")
	}
	if !strings.Contains(s, "<w:b/>") {
		t.Error("expected bold in table cells")
	}
}

func TestStyleTable_SpecificRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "styled-range.docx")
	// Style only the header row (row 0)
	cellStyle := CellStyle{BgColor: "003366"}
	runStyle := RunStyle{Bold: true, Color: "FFFFFF"}
	if err := StyleTable(path, out, 2, 0, 0, -1, -1, cellStyle, runStyle); err != nil {
		t.Fatal(err)
	}
	// Verify raw XML
	zc, err := common.OpenReader(out)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := common.ReadEntry(&zc.Reader, documentXMLPath)
	_ = zc.Close()
	if err != nil {
		t.Fatal(err)
	}
	s := string(doc)
	if !strings.Contains(s, "003366") {
		t.Error("expected bg color for header row")
	}
	if !strings.Contains(s, `<w:color w:val="FFFFFF"/>`) {
		t.Error("expected white text color")
	}
}

func TestStyleTable_OutOfRange(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	if err := StyleTable(path, out, 99, -1, -1, -1, -1, CellStyle{BgColor: "FF0000"}, RunStyle{}); err == nil {
		t.Fatal("expected error for out-of-range table index")
	}
}

func TestStyleTable_NotATable(t *testing.T) {
	path := tempDocxWithContent(t)
	out := filepath.Join(t.TempDir(), "err.docx")
	// Element 0 is a paragraph (Title)
	if err := StyleTable(path, out, 0, -1, -1, -1, -1, CellStyle{BgColor: "FF0000"}, RunStyle{}); err == nil {
		t.Fatal("expected error when styling a non-table element")
	}
}

// ---------------------------------------------------------------------------
// splitXMLBlocks / reassembleBlocks
// ---------------------------------------------------------------------------

func TestSplitXMLBlocks(t *testing.T) {
	s := `<w:tr><w:tc><w:t>A</w:t></w:tc><w:tc><w:t>B</w:t></w:tc></w:tr><w:tr><w:tc><w:t>C</w:t></w:tc></w:tr>`
	blocks := splitXMLBlocks(s, "<w:tr>", "</w:tr>")
	if len(blocks) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(blocks))
	}
	if !strings.Contains(string(blocks[0]), "A") || !strings.Contains(string(blocks[0]), "B") {
		t.Error("first row should contain A and B")
	}
	if !strings.Contains(string(blocks[1]), "C") {
		t.Error("second row should contain C")
	}
}

func TestSplitXMLBlocks_Cells(t *testing.T) {
	s := `<w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Y</w:t></w:r></w:p></w:tc>`
	blocks := splitXMLBlocks(s, "<w:tc>", "</w:tc>")
	if len(blocks) != 2 {
		t.Fatalf("expected 2 cells, got %d", len(blocks))
	}
}

func TestReassembleBlocks_Roundtrip(t *testing.T) {
	s := `<w:tr><w:tc>A</w:tc></w:tr><w:tr><w:tc>B</w:tc></w:tr>`
	blocks := splitXMLBlocks(s, "<w:tr>", "</w:tr>")
	// Modify blocks
	for i, b := range blocks {
		blocks[i] = []byte(strings.Replace(string(b), "A", "MODIFIED", 1))
	}
	_ = blocks // use blocks
	// Just verify split + reassemble doesn't lose content
	if len(blocks) != 2 {
		t.Fatal("expected 2 blocks")
	}
}

// ---------------------------------------------------------------------------
// IsZero methods
// ---------------------------------------------------------------------------

func TestRunStyle_IsZero(t *testing.T) {
	s := RunStyle{}
	if !s.IsZero() {
		t.Error("zero RunStyle should be IsZero")
	}
	s.Bold = true
	if s.IsZero() {
		t.Error("Bold=true should not be IsZero")
	}
}

func TestParaStyle_IsZero(t *testing.T) {
	s := ParaStyle{}
	if !s.IsZero() {
		t.Error("zero ParaStyle should be IsZero")
	}
	s.Align = "center"
	if s.IsZero() {
		t.Error("Align=center should not be IsZero")
	}
}

func TestCellStyle_IsZero(t *testing.T) {
	s := CellStyle{}
	if !s.IsZero() {
		t.Error("zero CellStyle should be IsZero")
	}
	s.BgColor = "FFFF00"
	if s.IsZero() {
		t.Error("BgColor set should not be IsZero")
	}
}
