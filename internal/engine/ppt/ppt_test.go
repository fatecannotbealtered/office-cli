package ppt

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// makePPTX creates a minimal .pptx with 3 slides for testing.
func makePPTX(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "test.pptx")

	f, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	zw := zip.NewWriter(f)
	defer func() { _ = zw.Close() }()

	// [Content_Types].xml
	writeEntry(t, zw, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)

	// Slide 1: title + bullet
	writeEntry(t, zw, "ppt/slides/slide1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sld xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld><spTree>
    <sp><nvSpPr><nvPr><ph type="title"/></nvPr></nvSpPr>
      <txBody><p><r><t>Slide One Title</t></r></p></txBody>
    </sp>
    <sp><nvSpPr><nvPr/></nvSpPr>
      <txBody><p><r><t>Bullet A</t></r></p><p><r><t>Bullet B</t></r></p></txBody>
    </sp>
  </spTree></cSld>
</sld>`)

	// Slide 2: title only
	writeEntry(t, zw, "ppt/slides/slide2.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sld xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld><spTree>
    <sp><nvSpPr><nvPr><ph type="ctrTitle"/></nvPr></nvSpPr>
      <txBody><p><r><t>Second Slide</t></r></p></txBody>
    </sp>
  </spTree></cSld>
</sld>`)

	// Slide 3: no title placeholder, just text
	writeEntry(t, zw, "ppt/slides/slide3.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sld xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">
  <cSld><spTree>
    <sp><nvSpPr><nvPr/></nvSpPr>
      <txBody><p><r><t>{{MILESTONE}}</t></r></p></txBody>
    </sp>
  </spTree></cSld>
</sld>`)

	// Minimal relationships
	writeEntry(t, zw, "_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`)

	return path
}

func writeEntry(t *testing.T, zw *zip.Writer, name, content string) {
	t.Helper()
	w, err := zw.Create(name)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Write([]byte(content)); err != nil {
		t.Fatal(err)
	}
}

func TestReadSlides(t *testing.T) {
	path := makePPTX(t)

	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 3 {
		t.Fatalf("expected 3 slides, got %d", len(slides))
	}

	// Slide 1
	if slides[0].Title != "Slide One Title" {
		t.Errorf("slide 1 title = %q, want %q", slides[0].Title, "Slide One Title")
	}
	if len(slides[0].Bullets) < 2 {
		t.Errorf("slide 1 bullets = %d, want >= 2", len(slides[0].Bullets))
	}

	// Slide 2
	if slides[1].Title != "Second Slide" {
		t.Errorf("slide 2 title = %q, want %q", slides[1].Title, "Second Slide")
	}

	// Slide 3 - no title placeholder, falls back to first text
	if slides[2].Title != "{{MILESTONE}}" {
		t.Errorf("slide 3 title = %q, want %q", slides[2].Title, "{{MILESTONE}}")
	}
}

func TestSlideCount(t *testing.T) {
	path := makePPTX(t)

	n, err := SlideCount(path)
	if err != nil {
		t.Fatalf("SlideCount: %v", err)
	}
	if n != 3 {
		t.Errorf("count = %d, want 3", n)
	}
}

func TestReadOutline(t *testing.T) {
	path := makePPTX(t)

	outline, err := ReadOutline(path)
	if err != nil {
		t.Fatalf("ReadOutline: %v", err)
	}
	if len(outline) != 3 {
		t.Fatalf("expected 3, got %d", len(outline))
	}
	if outline[0].Title != "Slide One Title" {
		t.Errorf("outline[0].Title = %q", outline[0].Title)
	}
}

func TestReplaceText(t *testing.T) {
	path := makePPTX(t)
	outPath := filepath.Join(t.TempDir(), "replaced.pptx")

	hits, err := ReplaceText(path, outPath, []Replacement{
		{Find: "{{MILESTONE}}", Replace: "2026Q3"},
	})
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if hits != 1 {
		t.Errorf("hits = %d, want 1", hits)
	}

	// Verify replacement
	slides, err := ReadSlides(outPath)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	found := false
	for _, s := range slides {
		if s.Title == "2026Q3" || s.Text == "2026Q3" {
			found = true
		}
	}
	if !found {
		t.Error("replacement not found in output")
	}
}

func TestReplaceTextEmpty(t *testing.T) {
	path := makePPTX(t)
	_, err := ReplaceText(path, "/tmp/out.pptx", nil)
	if err == nil {
		t.Error("expected error for empty replacements")
	}
}

func TestReadMeta(t *testing.T) {
	path := makePPTX(t)

	meta, err := ReadMeta(path)
	if err != nil {
		t.Fatalf("ReadMeta: %v", err)
	}
	if meta.Slides != 3 {
		t.Errorf("slides = %d, want 3", meta.Slides)
	}
	if meta.Path != path {
		t.Errorf("path = %q, want %q", meta.Path, path)
	}
	if meta.SizeBytes == 0 {
		t.Error("sizeBytes should be > 0")
	}
}

func TestSlideNum(t *testing.T) {
	cases := []struct {
		name, prefix string
		want         int
	}{
		{"ppt/slides/slide1.xml", "ppt/slides/slide", 1},
		{"ppt/slides/slide12.xml", "ppt/slides/slide", 12},
		{"ppt/slides/slide.xml", "ppt/slides/slide", 0},
		{"ppt/slides/other.xml", "ppt/slides/slide", 0},
	}
	for _, tc := range cases {
		got := slideNum(tc.name, tc.prefix)
		if got != tc.want {
			t.Errorf("slideNum(%q, %q) = %d, want %d", tc.name, tc.prefix, got, tc.want)
		}
	}
}

// helper: read a pptx zip entry as string
func readZipEntry(t *testing.T, path, name string) string {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("reading %s: %v", path, err)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("opening zip %s: %v", path, err)
	}
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("opening entry %s: %v", name, err)
			}
			defer func() { _ = rc.Close() }()
			var buf bytes.Buffer
			if _, err := buf.ReadFrom(rc); err != nil {
				t.Fatalf("reading entry %s: %v", name, err)
			}
			return buf.String()
		}
	}
	t.Fatalf("entry %s not found in %s", name, path)
	return ""
}

// helper: count entries matching a prefix in a pptx zip
func countZipEntries(t *testing.T, path, prefix string) int {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("reading %s: %v", path, err)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("opening zip %s: %v", path, err)
	}
	n := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, prefix) {
			n++
		}
	}
	return n
}

func TestCreate(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "created.pptx")

	if err := Create(path, "My Title", "Alice"); err != nil {
		t.Fatalf("Create: %v", err)
	}

	// Verify the file exists and has key XML parts
	ct := readZipEntry(t, path, "[Content_Types].xml")
	if !strings.Contains(ct, "presentationml.slide+xml") {
		t.Error("Content_Types missing slide override")
	}

	pres := readZipEntry(t, path, "ppt/presentation.xml")
	if !strings.Contains(pres, "sldIdLst") {
		t.Error("presentation.xml missing sldIdLst")
	}

	// Should have exactly 1 slide
	count := countZipEntries(t, path, "ppt/slides/slide")
	if count != 1 {
		t.Errorf("expected 1 slide file, got %d", count)
	}

	// Read back and verify
	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
}

func TestAddSlide(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "add.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}

	// Add a slide with title and bullets
	if err := AddSlide(path, path, "Roadmap", []string{"Q1: done", "Q2: WIP"}); err != nil {
		t.Fatalf("AddSlide: %v", err)
	}

	// Should now have 2 slides
	count := countZipEntries(t, path, "ppt/slides/slide")
	if count != 2 {
		t.Errorf("expected 2 slide files, got %d", count)
	}

	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 2 {
		t.Fatalf("expected 2 slides, got %d", len(slides))
	}

	// Second slide should have our title
	if slides[1].Title != "Roadmap" {
		t.Errorf("slide 2 title = %q, want %q", slides[1].Title, "Roadmap")
	}
}

func TestSetSlideContent(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "set.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}

	// Overwrite slide 1
	if err := SetSlideContent(path, path, 1, "New Title", []string{"Point A", "Point B", "Point C"}); err != nil {
		t.Fatalf("SetSlideContent: %v", err)
	}

	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
	if slides[0].Title != "New Title" {
		t.Errorf("title = %q, want %q", slides[0].Title, "New Title")
	}
	if len(slides[0].Bullets) < 3 {
		t.Errorf("expected >= 3 bullets, got %d", len(slides[0].Bullets))
	}
}

func TestSetNotes(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "notes.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}

	if err := SetNotes(path, path, 1, "Remember to smile!"); err != nil {
		t.Fatalf("SetNotes: %v", err)
	}

	// Notes file should exist
	ct := readZipEntry(t, path, "[Content_Types].xml")
	if !strings.Contains(ct, "notesSlide1.xml") {
		t.Error("Content_Types missing notesSlide1.xml override")
	}

	notesXML := readZipEntry(t, path, "ppt/notesSlides/notesSlide1.xml")
	if !strings.Contains(notesXML, "Remember to smile!") {
		t.Error("notesSlide1.xml missing notes text")
	}

	// Slide's .rels should reference the notes slide
	relsFile := "ppt/slides/_rels/slide1.xml.rels"
	rels := readZipEntry(t, path, relsFile)
	if !strings.Contains(rels, "notesSlide") {
		t.Errorf("slide 1 .rels missing notesSlide relationship: %s", rels)
	}
	if !strings.Contains(rels, "notesSlide1.xml") {
		t.Errorf("slide 1 .rels missing notesSlide1.xml target: %s", rels)
	}
}

func TestDeleteSlide(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "del.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Second", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Third", nil); err != nil {
		t.Fatal(err)
	}

	// Delete slide 2
	if err := DeleteSlide(path, path, 2); err != nil {
		t.Fatalf("DeleteSlide: %v", err)
	}

	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 2 {
		t.Fatalf("expected 2 slides after delete, got %d", len(slides))
	}
	// After delete, Index is presentation position (1,2), not file number.
	// Verify by title: slide file 1 ("") and file 3 ("Third") should remain.
	if len(slides) != 2 {
		t.Fatalf("expected 2 slides, got %d", len(slides))
	}
	// Check that slide file 1 (slide1.xml) and file 3 (slide3.xml) are present
	foundFile1, foundFile3 := false, false
	for _, s := range slides {
		if strings.HasSuffix(s.File, "slide1.xml") {
			foundFile1 = true
		}
		if strings.HasSuffix(s.File, "slide3.xml") {
			foundFile3 = true
		}
	}
	if !foundFile1 || !foundFile3 {
		t.Errorf("expected slide1.xml and slide3.xml to remain, found1=%v found3=%v", foundFile1, foundFile3)
	}
}

func TestReorderSlides(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "reorder.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}
	// Set slide 1 content so we can identify it
	if err := SetSlideContent(path, path, 1, "First", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Second", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Third", nil); err != nil {
		t.Fatal(err)
	}

	// Reorder: 3, 1, 2
	if err := ReorderSlides(path, path, "3,1,2"); err != nil {
		t.Fatalf("ReorderSlides: %v", err)
	}

	// Verify the sldIdLst order by reading the raw presentation.xml
	// Reorder only changes the sldIdLst, not the slide file numbering.
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	presData, err := func() ([]byte, error) {
		for _, f := range zr.File {
			if f.Name == "ppt/presentation.xml" {
				rc, _ := f.Open()
				defer func() { _ = rc.Close() }()
				var buf bytes.Buffer
				_, _ = buf.ReadFrom(rc)
				return buf.Bytes(), nil
			}
		}
		return nil, os.ErrNotExist
	}()
	if err != nil {
		t.Fatal(err)
	}

	entries, _, err := parseSldIDLst(presData)
	if err != nil {
		t.Fatalf("parseSldIDLst: %v", err)
	}
	if len(entries) != 3 {
		t.Fatalf("expected 3 sldId entries, got %d", len(entries))
	}

	// The rels mapping: slide 1 → rId1, slide 2 → rId2, slide 3 → rId3
	// After reorder 3,1,2: the sldIdLst should reference slides 3, 1, 2
	// We verify by checking the rId order matches slide 3, slide 1, slide 2
	relsStr := readZipEntry(t, path, "ppt/_rels/presentation.xml.rels")
	var rels xmlRels
	_ = xml.Unmarshal([]byte(relsStr), &rels)
	rIDToSlide := map[string]int{}
	slideRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	for _, r := range rels.Entries {
		if r.Type == slideRelType {
			num := slideNum("ppt/"+r.Target, "ppt/slides/slide")
			rIDToSlide[r.ID] = num
		}
	}

	var slideOrder []int
	for _, e := range entries {
		slideOrder = append(slideOrder, rIDToSlide[e.RId])
	}
	// Expected order: slide 3 first, then 1, then 2
	if slideOrder[0] != 3 || slideOrder[1] != 1 || slideOrder[2] != 2 {
		t.Errorf("sldIdLst slide order = %v, want [3 1 2]", slideOrder)
	}
}

func TestAddImage(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "img.pptx")
	imgPath := filepath.Join(dir, "test.png")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}

	// Create a minimal valid 1x1 PNG (67 bytes)
	pngData := []byte{
		0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, // PNG signature
		0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
		0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, // 8-bit RGB
		0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
		0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
		0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33,
		0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, // IEND chunk
		0xae, 0x42, 0x60, 0x82,
	}
	if err := os.WriteFile(imgPath, pngData, 0644); err != nil {
		t.Fatal(err)
	}

	if err := AddImage(path, path, 1, imgPath, 0, 0); err != nil {
		t.Fatalf("AddImage: %v", err)
	}

	// Verify media file was added
	mediaCount := countZipEntries(t, path, "ppt/media/")
	if mediaCount < 1 {
		t.Error("expected at least 1 media file after AddImage")
	}

	// Verify slide XML contains <p:pic> with <a:blip>
	slideXML := readZipEntry(t, path, "ppt/slides/slide1.xml")
	if !strings.Contains(slideXML, "<p:pic>") {
		t.Error("slide XML missing <p:pic> element")
	}
	if !strings.Contains(slideXML, "<a:blip") {
		t.Error("slide XML missing <a:blip> element")
	}
	if !strings.Contains(slideXML, "r:embed=") {
		t.Error("slide XML missing r:embed attribute")
	}
}

func TestBuild(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "built.pptx")

	spec := BuildSpec{
		Title:  "Q2 Review",
		Author: "Agent",
		Slides: []SlideSpec{
			{Title: "Overview", Bullets: []string{"Revenue up 20%", "Users doubled"}},
			{Title: "Roadmap", Bullets: []string{"Q3: launch", "Q4: scale"}},
			{Title: "Q&A", Notes: "Remember to thank the team"},
		},
	}

	if err := Build(path, spec); err != nil {
		t.Fatalf("Build: %v", err)
	}

	// Verify slide count
	count, err := SlideCount(path)
	if err != nil {
		t.Fatalf("SlideCount: %v", err)
	}
	if count != 3 {
		t.Errorf("expected 3 slides, got %d", count)
	}

	// Verify content
	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 3 {
		t.Fatalf("expected 3 slides, got %d", len(slides))
	}
	if slides[0].Title != "Overview" {
		t.Errorf("slide 1 title = %q, want %q", slides[0].Title, "Overview")
	}
	if slides[1].Title != "Roadmap" {
		t.Errorf("slide 2 title = %q, want %q", slides[1].Title, "Roadmap")
	}
	if slides[2].Title != "Q&A" {
		t.Errorf("slide 3 title = %q, want %q", slides[2].Title, "Q&A")
	}

	// Verify notes for slide 3
	notesXML := readZipEntry(t, path, "ppt/notesSlides/notesSlide3.xml")
	if !strings.Contains(notesXML, "Remember to thank the team") {
		t.Error("notesSlide3.xml missing notes text")
	}

	// Verify metadata
	meta, err := ReadMeta(path)
	if err != nil {
		t.Fatalf("ReadMeta: %v", err)
	}
	if meta.Title != "Q2 Review" {
		t.Errorf("meta title = %q, want %q", meta.Title, "Q2 Review")
	}
	if meta.Author != "Agent" {
		t.Errorf("meta author = %q, want %q", meta.Author, "Agent")
	}
}

// TestSlideRelMap verifies that parseSlideRelMap correctly identifies rIds.
func TestSlideRelMap(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "relmap.pptx")

	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}
	// Add slides to create non-trivial rId mapping
	if err := AddSlide(path, path, "B", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "C", nil); err != nil {
		t.Fatal(err)
	}

	// Delete slide 2 — this should leave a gap in rIds
	if err := DeleteSlide(path, path, 2); err != nil {
		t.Fatal(err)
	}

	// Add a new slide — its rId should be max+1, not reuse the deleted one
	if err := AddSlide(path, path, "D", nil); err != nil {
		t.Fatal(err)
	}

	// Read back and verify we still have 3 slides
	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 3 {
		t.Fatalf("expected 3 slides, got %d", len(slides))
	}

	// Verify presentation.xml sldIdLst can be parsed
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	presData, err := func() ([]byte, error) {
		for _, f := range zr.File {
			if f.Name == "ppt/presentation.xml" {
				rc, _ := f.Open()
				defer func() { _ = rc.Close() }()
				var buf bytes.Buffer
				_, _ = buf.ReadFrom(rc)
				return buf.Bytes(), nil
			}
		}
		return nil, os.ErrNotExist
	}()
	if err != nil {
		t.Fatal(err)
	}
	// Use token-based parser (handles namespace attributes correctly)
	entries, maxSldID, err := parseSldIDLst(presData)
	if err != nil {
		t.Fatalf("parseSldIDLst: %v", err)
	}
	if len(entries) != 3 {
		t.Errorf("expected 3 sldId entries, got %d", len(entries))
	}
	if maxSldID < 256 {
		t.Errorf("maxSldID = %d, want >= 256", maxSldID)
	}
}

func TestReadSlidesOrder(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "order.pptx")

	// Create deck with 3 slides
	if err := Create(path, "", ""); err != nil {
		t.Fatal(err)
	}
	if err := SetSlideContent(path, path, 1, "First", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Second", nil); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(path, path, "Third", nil); err != nil {
		t.Fatal(err)
	}

	// Before reorder: should be 1,2,3
	slides, err := ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if slides[0].Title != "First" || slides[1].Title != "Second" || slides[2].Title != "Third" {
		t.Errorf("before reorder: got [%s, %s, %s], want [First, Second, Third]",
			slides[0].Title, slides[1].Title, slides[2].Title)
	}

	// Reorder: 3,1,2
	if err := ReorderSlides(path, path, "3,1,2"); err != nil {
		t.Fatal(err)
	}

	// After reorder: should be Third, First, Second
	slides, err = ReadSlides(path)
	if err != nil {
		t.Fatalf("ReadSlides: %v", err)
	}
	if len(slides) != 3 {
		t.Fatalf("expected 3 slides, got %d", len(slides))
	}
	if slides[0].Title != "Third" {
		t.Errorf("slide 1 title = %q, want %q", slides[0].Title, "Third")
	}
	if slides[1].Title != "First" {
		t.Errorf("slide 2 title = %q, want %q", slides[1].Title, "First")
	}
	if slides[2].Title != "Second" {
		t.Errorf("slide 3 title = %q, want %q", slides[2].Title, "Second")
	}

	// Index should be presentation position (1,2,3), not file number
	if slides[0].Index != 1 || slides[1].Index != 2 || slides[2].Index != 3 {
		t.Errorf("indices = [%d, %d, %d], want [1, 2, 3]",
			slides[0].Index, slides[1].Index, slides[2].Index)
	}
}

func TestBuildWithImage(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "buildimg.pptx")
	imgPath := filepath.Join(dir, "test.png")

	// Create a minimal valid 1x1 PNG
	pngData := []byte{
		0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
		0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde,
		0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54,
		0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
		0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33,
		0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
		0xae, 0x42, 0x60, 0x82,
	}
	if err := os.WriteFile(imgPath, pngData, 0644); err != nil {
		t.Fatal(err)
	}

	spec := BuildSpec{
		Slides: []SlideSpec{
			{Title: "With Image", Bullets: []string{"See below"}, Image: imgPath},
			{Title: "No Image", Bullets: []string{"Just text"}},
		},
	}

	if err := Build(path, spec); err != nil {
		t.Fatalf("Build: %v", err)
	}

	// Verify slide 1 has <p:pic> (image embedded)
	slide1 := readZipEntry(t, path, "ppt/slides/slide1.xml")
	if !strings.Contains(slide1, "<p:pic>") {
		t.Error("slide 1 missing <p:pic> element")
	}
	if !strings.Contains(slide1, "<a:blip") {
		t.Error("slide 1 missing <a:blip> element")
	}

	// Verify slide 2 does NOT have <p:pic>
	slide2 := readZipEntry(t, path, "ppt/slides/slide2.xml")
	if strings.Contains(slide2, "<p:pic>") {
		t.Error("slide 2 should not have <p:pic> element")
	}

	// Verify media file exists
	mediaCount := countZipEntries(t, path, "ppt/media/")
	if mediaCount < 1 {
		t.Error("expected at least 1 media file")
	}
}

func TestBuildFromTemplate(t *testing.T) {
	dir := t.TempDir()
	templatePath := filepath.Join(dir, "template.pptx")
	outPath := filepath.Join(dir, "from_template.pptx")

	// Create a "template" with 2 slides
	if err := Create(templatePath, "Original Title", "Original Author"); err != nil {
		t.Fatal(err)
	}
	if err := AddSlide(templatePath, templatePath, "Template Slide 2", nil); err != nil {
		t.Fatal(err)
	}

	// Verify template has 2 slides
	count, _ := SlideCount(templatePath)
	if count != 2 {
		t.Fatalf("template has %d slides, want 2", count)
	}

	// Build from template with 3 slides
	spec := BuildSpec{
		Title:  "New Title",
		Author: "New Author",
		Slides: []SlideSpec{
			{Title: "Overview", Bullets: []string{"Point A"}},
			{Title: "Details", Bullets: []string{"Point B"}},
			{Title: "Summary", Notes: "Wrap up here"},
		},
	}

	if err := BuildFromTemplate(templatePath, outPath, spec); err != nil {
		t.Fatalf("BuildFromTemplate: %v", err)
	}

	// Verify slide count
	count, err := SlideCount(outPath)
	if err != nil {
		t.Fatal(err)
	}
	if count != 3 {
		t.Errorf("expected 3 slides, got %d", count)
	}

	// Verify content
	slides, err := ReadSlides(outPath)
	if err != nil {
		t.Fatal(err)
	}
	if slides[0].Title != "Overview" {
		t.Errorf("slide 1 title = %q, want %q", slides[0].Title, "Overview")
	}
	if slides[1].Title != "Details" {
		t.Errorf("slide 2 title = %q, want %q", slides[1].Title, "Details")
	}
	if slides[2].Title != "Summary" {
		t.Errorf("slide 3 title = %q, want %q", slides[2].Title, "Summary")
	}

	// Verify notes
	notesXML := readZipEntry(t, outPath, "ppt/notesSlides/notesSlide3.xml")
	if !strings.Contains(notesXML, "Wrap up here") {
		t.Error("notesSlide3.xml missing notes text")
	}

	// Verify metadata was updated
	meta, err := ReadMeta(outPath)
	if err != nil {
		t.Fatal(err)
	}
	if meta.Title != "New Title" {
		t.Errorf("meta title = %q, want %q", meta.Title, "New Title")
	}
	if meta.Author != "New Author" {
		t.Errorf("meta author = %q, want %q", meta.Author, "New Author")
	}

	// Verify the template file was not modified
	origMeta, _ := ReadMeta(templatePath)
	if origMeta.Title != "Original Title" {
		t.Errorf("template title changed to %q", origMeta.Title)
	}
}

// ---------------------------------------------------------------------------
// Phase 3: Layout / Style / Shape tests
// ---------------------------------------------------------------------------

// makePPTXWithStyles creates a .pptx with styled shapes for layout testing.
func makePPTXWithStyles(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "styled.pptx")

	f, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	zw := zip.NewWriter(f)
	defer func() { _ = zw.Close() }()

	writeEntry(t, zw, "[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`)

	// Slide 1: title shape with font size + bold + color, body shape with bullets
	writeEntry(t, zw, "ppt/slides/slide1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1371600"/></a:xfrm>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p>
        <a:r><a:rPr lang="en-US" sz="3600" b="1"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:rPr><a:t>Q2 Review</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Content Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="5029200"/></a:xfrm>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p>
        <a:r><a:rPr lang="en-US" sz="1800"/><a:t>Revenue up 20%</a:t></a:r>
      </a:p><a:p>
        <a:r><a:rPr lang="en-US" sz="1800" b="1"/><a:t>Key milestone</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</sld>`)

	// Slide 2: no placeholder, just a plain text box
	writeEntry(t, zw, "ppt/slides/slide2.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="10" name="TextBox 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="100000" y="100000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:p>
        <a:pPr algn="ctr"/>
        <a:r><a:rPr lang="en-US" sz="2400" i="1" u="sng"><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill></a:rPr><a:t>Centered italic</a:t></a:r>
      </a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</sld>`)

	writeEntry(t, zw, "_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`)

	return path
}

func TestReadSlideLayout(t *testing.T) {
	path := makePPTXWithStyles(t)

	shapes, err := ReadSlideLayout(path, 1)
	if err != nil {
		t.Fatalf("ReadSlideLayout: %v", err)
	}
	if len(shapes) != 2 {
		t.Fatalf("expected 2 shapes on slide 1, got %d", len(shapes))
	}

	// Shape 0: title
	s0 := shapes[0]
	if s0.Type != "sp" {
		t.Errorf("shape 0 type = %q, want sp", s0.Type)
	}
	if s0.Name != "Title 1" {
		t.Errorf("shape 0 name = %q, want Title 1", s0.Name)
	}
	if s0.Ph != "title" {
		t.Errorf("shape 0 ph = %q, want title", s0.Ph)
	}
	if s0.X != 0 || s0.Y != 0 || s0.W != 9144000 || s0.H != 1371600 {
		t.Errorf("shape 0 pos = (%d,%d,%d,%d), want (0,0,9144000,1371600)", s0.X, s0.Y, s0.W, s0.H)
	}
	if s0.Text != "Q2 Review" {
		t.Errorf("shape 0 text = %q, want Q2 Review", s0.Text)
	}
	if len(s0.Paragraphs) != 1 {
		t.Fatalf("shape 0 paragraphs = %d, want 1", len(s0.Paragraphs))
	}
	r0 := s0.Paragraphs[0].Runs[0]
	if r0.FontSize != 3600 {
		t.Errorf("shape 0 run fontSize = %d, want 3600", r0.FontSize)
	}
	if !r0.Bold {
		t.Error("shape 0 run should be bold")
	}
	if r0.Color != "003366" {
		t.Errorf("shape 0 run color = %q, want 003366", r0.Color)
	}

	// Shape 1: body with 2 paragraphs
	s1 := shapes[1]
	if s1.Ph != "body" {
		t.Errorf("shape 1 ph = %q, want body", s1.Ph)
	}
	if s1.X != 457200 || s1.Y != 1600200 {
		t.Errorf("shape 1 pos = (%d,%d), want (457200,1600200)", s1.X, s1.Y)
	}
	if len(s1.Paragraphs) != 2 {
		t.Fatalf("shape 1 paragraphs = %d, want 2", len(s1.Paragraphs))
	}
	if s1.Text != "Revenue up 20%\nKey milestone" {
		t.Errorf("shape 1 text = %q", s1.Text)
	}
	if !s1.Paragraphs[1].Runs[0].Bold {
		t.Error("shape 1 para 1 run should be bold")
	}
}

func TestReadSlideLayoutSlide2(t *testing.T) {
	path := makePPTXWithStyles(t)

	shapes, err := ReadSlideLayout(path, 2)
	if err != nil {
		t.Fatalf("ReadSlideLayout: %v", err)
	}
	if len(shapes) != 1 {
		t.Fatalf("expected 1 shape on slide 2, got %d", len(shapes))
	}

	s := shapes[0]
	if s.Name != "TextBox 1" {
		t.Errorf("shape name = %q, want TextBox 1", s.Name)
	}
	if s.Ph != "" {
		t.Errorf("shape ph = %q, want empty", s.Ph)
	}
	if s.X != 100000 || s.Y != 100000 || s.W != 4000000 || s.H != 2000000 {
		t.Errorf("shape pos = (%d,%d,%d,%d)", s.X, s.Y, s.W, s.H)
	}
	if len(s.Paragraphs) != 1 {
		t.Fatalf("paragraphs = %d, want 1", len(s.Paragraphs))
	}
	if s.Paragraphs[0].Align != "ctr" {
		t.Errorf("align = %q, want ctr", s.Paragraphs[0].Align)
	}
	ri := s.Paragraphs[0].Runs[0]
	if ri.FontSize != 2400 {
		t.Errorf("fontSize = %d, want 2400", ri.FontSize)
	}
	if !ri.Italic {
		t.Error("should be italic")
	}
	if !ri.Underline {
		t.Error("should be underlined")
	}
	if ri.Color != "CC0000" {
		t.Errorf("color = %q, want CC0000", ri.Color)
	}
}

func TestSetShapeStyle(t *testing.T) {
	path := makePPTXWithStyles(t)
	outPath := filepath.Join(t.TempDir(), "out.pptx")

	bold := true
	fontSize := 4800
	color := "FF0000"
	align := "ctr"
	opts := StyleOptions{
		FontSize: &fontSize,
		Bold:     &bold,
		Color:    &color,
		Align:    &align,
	}

	if err := SetShapeStyle(path, outPath, 1, 0, opts); err != nil {
		t.Fatalf("SetShapeStyle: %v", err)
	}

	// Read back and verify
	shapes, err := ReadSlideLayout(outPath, 1)
	if err != nil {
		t.Fatal(err)
	}
	s0 := shapes[0]
	if len(s0.Paragraphs) != 1 {
		t.Fatalf("paragraphs = %d, want 1", len(s0.Paragraphs))
	}
	r0 := s0.Paragraphs[0].Runs[0]
	if r0.FontSize != 4800 {
		t.Errorf("fontSize = %d, want 4800", r0.FontSize)
	}
	if !r0.Bold {
		t.Error("should be bold")
	}
	if r0.Color != "FF0000" {
		t.Errorf("color = %q, want FF0000", r0.Color)
	}
	if s0.Paragraphs[0].Align != "ctr" {
		t.Errorf("align = %q, want ctr", s0.Paragraphs[0].Align)
	}
}

func TestAddShape(t *testing.T) {
	path := makePPTX(t)
	outPath := filepath.Join(t.TempDir(), "shaped.pptx")

	spec := ShapeSpec{
		Type:     "rect",
		X:        500000,
		Y:        200000,
		W:        4000000,
		H:        1000000,
		Text:     "Hello World",
		FontSize: 2400,
		Bold:     true,
		Color:    "336699",
		Fill:     "E8E8E8",
	}

	if err := AddShape(path, outPath, 1, spec); err != nil {
		t.Fatalf("AddShape: %v", err)
	}

	// Read back and verify the shape appears in layout
	shapes, err := ReadSlideLayout(outPath, 1)
	if err != nil {
		t.Fatal(err)
	}
	// Original slide has 2 shapes (title + body), new shape is appended
	found := false
	for _, s := range shapes {
		if s.Type == "sp" && s.X == 500000 && s.Y == 200000 && s.W == 4000000 {
			found = true
			if s.Text != "Hello World" {
				t.Errorf("shape text = %q, want Hello World", s.Text)
			}
			if len(s.Paragraphs) > 0 && len(s.Paragraphs[0].Runs) > 0 {
				r := s.Paragraphs[0].Runs[0]
				if r.FontSize != 2400 {
					t.Errorf("fontSize = %d, want 2400", r.FontSize)
				}
				if !r.Bold {
					t.Error("should be bold")
				}
				if r.Color != "336699" {
					t.Errorf("color = %q, want 336699", r.Color)
				}
			}
			break
		}
	}
	if !found {
		t.Error("added shape not found in layout output")
	}
}

func TestAddShapeLine(t *testing.T) {
	path := makePPTX(t)
	outPath := filepath.Join(t.TempDir(), "line.pptx")

	spec := ShapeSpec{
		Type: "line",
		X:    100000,
		Y:    300000,
		W:    5000000,
		H:    0,
		Line: "000000",
	}

	if err := AddShape(path, outPath, 1, spec); err != nil {
		t.Fatalf("AddShape (line): %v", err)
	}

	// Verify the shape was added by reading raw XML
	xmlStr := readZipEntry(t, outPath, "ppt/slides/slide1.xml")
	if !strings.Contains(xmlStr, `prst="line"`) {
		t.Error("slide XML should contain prst=\"line\"")
	}
}
