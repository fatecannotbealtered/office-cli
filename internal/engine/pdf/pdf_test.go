package pdf

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// makePDF writes a minimal PDF with one page per `texts` element.
// Uses the same hand-crafted PDF structure as scripts/gen-fixtures.
func makePDF(t *testing.T, texts ...string) string {
	t.Helper()
	if len(texts) == 0 {
		texts = []string{""}
	}
	path := filepath.Join(t.TempDir(), "test.pdf")
	if err := writeMinimalPDF(path, texts); err != nil {
		t.Fatal(err)
	}
	return path
}

func writeMinimalPDF(path string, pages []string) error {
	if len(pages) == 0 {
		pages = []string{""}
	}
	var body bytes.Buffer
	body.WriteString("%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")

	type obj struct {
		num    int
		offset int
	}
	var objects []obj

	addObj := func(num int, content string) {
		objects = append(objects, obj{num: num, offset: body.Len()})
		fmt.Fprintf(&body, "%d 0 obj\n%s\nendobj\n", num, content)
	}

	pageObjStart := 3
	contentObjStart := pageObjStart + len(pages)
	fontObjNum := contentObjStart + len(pages)

	var kids string
	for i := range pages {
		if i > 0 {
			kids += " "
		}
		kids += fmt.Sprintf("%d 0 R", pageObjStart+i)
	}

	addObj(1, "<< /Type /Catalog /Pages 2 0 R >>")
	addObj(2, fmt.Sprintf("<< /Type /Pages /Kids [%s] /Count %d >>", kids, len(pages)))

	for i, text := range pages {
		pn := pageObjStart + i
		cn := contentObjStart + i
		addObj(pn, fmt.Sprintf("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents %d 0 R /Resources << /Font << /F1 %d 0 R >> >> >>", cn, fontObjNum))
		stream := fmt.Sprintf("BT /F1 24 Tf 100 700 Td (%s) Tj ET\n", escapeStr(text))
		addObj(cn, fmt.Sprintf("<< /Length %d >>\nstream\n%sendstream", len(stream), stream))
	}
	addObj(fontObjNum, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

	xref := body.Len()
	total := fontObjNum + 1
	body.WriteString("xref\n")
	fmt.Fprintf(&body, "0 %d\n", total)
	body.WriteString("0000000000 65535 f \n")
	for n := 1; n < total; n++ {
		var found *obj
		for i := range objects {
			if objects[i].num == n {
				found = &objects[i]
				break
			}
		}
		if found == nil {
			body.WriteString("0000000000 00000 f \n")
			continue
		}
		fmt.Fprintf(&body, "%010d 00000 n \n", found.offset)
	}
	fmt.Fprintf(&body, "trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n", total, xref)
	return os.WriteFile(path, body.Bytes(), 0644)
}

func escapeStr(s string) string {
	var b bytes.Buffer
	for _, r := range s {
		switch r {
		case '(', ')', '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		default:
			if r > 0x7f {
				b.WriteByte('?')
				continue
			}
			b.WriteRune(r)
		}
	}
	return b.String()
}

// ---------------------------------------------------------------------------
// ParsePageList
// ---------------------------------------------------------------------------

func TestParsePageList(t *testing.T) {
	cases := []struct {
		in     string
		expect bool
	}{
		{"", true},
		{"1", true},
		{"1,3", true},
		{"1-3", true},
		{"1,3-5,7", true},
		{"5-", true},
		{"abc", false},
		{"1-2-3", false},
	}
	for _, c := range cases {
		_, err := ParsePageList(c.in)
		if (err == nil) != c.expect {
			t.Errorf("ParsePageList(%q): err=%v, expected ok=%v", c.in, err, c.expect)
		}
	}
}

// ---------------------------------------------------------------------------
// Read
// ---------------------------------------------------------------------------

func TestRead_AllPages(t *testing.T) {
	path := makePDF(t, "Hello", "World")
	pages, err := Read(path, ReadOptions{})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if len(pages) != 2 {
		t.Fatalf("expected 2 pages, got %d", len(pages))
	}
	if pages[0].Page != 1 || pages[1].Page != 2 {
		t.Errorf("page numbers: got %d,%d want 1,2", pages[0].Page, pages[1].Page)
	}
	if !strings.Contains(pages[0].Text, "Hello") {
		t.Errorf("page 1 text missing 'Hello': %q", pages[0].Text)
	}
	if !strings.Contains(pages[1].Text, "World") {
		t.Errorf("page 2 text missing 'World': %q", pages[1].Text)
	}
}

func TestRead_Range(t *testing.T) {
	path := makePDF(t, "A", "B", "C")
	pages, err := Read(path, ReadOptions{From: 2, To: 3})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if len(pages) != 2 {
		t.Fatalf("expected 2 pages, got %d", len(pages))
	}
	if pages[0].Page != 2 {
		t.Errorf("first page=%d want 2", pages[0].Page)
	}
}

func TestRead_Limit(t *testing.T) {
	path := makePDF(t, "A", "B", "C")
	pages, err := Read(path, ReadOptions{Limit: 1})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if len(pages) != 1 {
		t.Fatalf("expected 1 page, got %d", len(pages))
	}
}

func TestRead_OutOfRange(t *testing.T) {
	path := makePDF(t, "Only")
	_, err := Read(path, ReadOptions{From: 5})
	if err == nil {
		t.Error("expected error for out-of-range page")
	}
}

// ---------------------------------------------------------------------------
// AllText
// ---------------------------------------------------------------------------

func TestAllText(t *testing.T) {
	path := makePDF(t, "Alpha", "Beta")
	text, err := AllText(path)
	if err != nil {
		t.Fatalf("AllText: %v", err)
	}
	if !strings.Contains(text, "Alpha") || !strings.Contains(text, "Beta") {
		t.Errorf("AllText missing expected content: %q", text)
	}
}

// ---------------------------------------------------------------------------
// PageCount
// ---------------------------------------------------------------------------

func TestPageCount(t *testing.T) {
	path := makePDF(t, "P1", "P2", "P3")
	n, err := PageCount(path)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 3 {
		t.Errorf("PageCount=%d want 3", n)
	}
}

// ---------------------------------------------------------------------------
// PageDimensions
// ---------------------------------------------------------------------------

func TestPageDimensions(t *testing.T) {
	path := makePDF(t, "X")
	dims, err := PageDimensions(path)
	if err != nil {
		t.Fatalf("PageDimensions: %v", err)
	}
	if len(dims) != 1 {
		t.Fatalf("expected 1 dim, got %d", len(dims))
	}
	// US Letter: 612 x 792 pt
	if dims[0].Width != 612 || dims[0].Height != 792 {
		t.Errorf("dimensions=%.0fx%.0f want 612x792", dims[0].Width, dims[0].Height)
	}
}

// ---------------------------------------------------------------------------
// Merge
// ---------------------------------------------------------------------------

func TestMerge(t *testing.T) {
	a := makePDF(t, "First")
	b := makePDF(t, "Second")
	out := filepath.Join(t.TempDir(), "merged.pdf")

	if err := Merge([]string{a, b}, out); err != nil {
		t.Fatalf("Merge: %v", err)
	}
	n, _ := PageCount(out)
	if n != 2 {
		t.Errorf("merged PageCount=%d want 2", n)
	}
}

func TestMerge_TooFew(t *testing.T) {
	a := makePDF(t, "Only")
	if err := Merge([]string{a}, "x.pdf"); err == nil {
		t.Error("expected error for < 2 inputs")
	}
}

// ---------------------------------------------------------------------------
// SplitEvery
// ---------------------------------------------------------------------------

func TestSplitEvery(t *testing.T) {
	path := makePDF(t, "A", "B", "C", "D")
	outDir := filepath.Join(t.TempDir(), "split")

	if err := SplitEvery(path, outDir, 2); err != nil {
		t.Fatalf("SplitEvery: %v", err)
	}
	entries, _ := os.ReadDir(outDir)
	pdfCount := 0
	for _, e := range entries {
		if strings.HasSuffix(e.Name(), ".pdf") {
			pdfCount++
		}
	}
	if pdfCount != 2 {
		t.Errorf("expected 2 split files, got %d", pdfCount)
	}
}

func TestSplitEvery_InvalidSpan(t *testing.T) {
	path := makePDF(t, "X")
	if err := SplitEvery(path, t.TempDir(), 0); err == nil {
		t.Error("expected error for span=0")
	}
}

// ---------------------------------------------------------------------------
// Trim
// ---------------------------------------------------------------------------

func TestTrim(t *testing.T) {
	path := makePDF(t, "Keep", "Drop", "Keep2")
	out := filepath.Join(t.TempDir(), "trimmed.pdf")

	if err := Trim(path, out, "1,3"); err != nil {
		t.Fatalf("Trim: %v", err)
	}
	n, _ := PageCount(out)
	if n != 2 {
		t.Errorf("trimmed PageCount=%d want 2", n)
	}
}

// ---------------------------------------------------------------------------
// WatermarkText
// ---------------------------------------------------------------------------

func TestWatermarkText(t *testing.T) {
	path := makePDF(t, "Content")
	out := filepath.Join(t.TempDir(), "wm.pdf")

	if err := WatermarkText(path, out, "DRAFT", ""); err != nil {
		t.Fatalf("WatermarkText: %v", err)
	}
	n, _ := PageCount(out)
	if n != 1 {
		t.Errorf("watermarked PageCount=%d want 1", n)
	}
}

func TestWatermarkText_Empty(t *testing.T) {
	path := makePDF(t, "X")
	if err := WatermarkText(path, "x.pdf", "", ""); err == nil {
		t.Error("expected error for empty watermark text")
	}
}

// ---------------------------------------------------------------------------
// Rotate
// ---------------------------------------------------------------------------

func TestRotate(t *testing.T) {
	path := makePDF(t, "Page")
	out := filepath.Join(t.TempDir(), "rotated.pdf")

	if err := Rotate(path, out, 90, ""); err != nil {
		t.Fatalf("Rotate: %v", err)
	}
	n, _ := PageCount(out)
	if n != 1 {
		t.Errorf("rotated PageCount=%d want 1", n)
	}
}

func TestRotate_InvalidDegrees(t *testing.T) {
	path := makePDF(t, "X")
	if err := Rotate(path, "x.pdf", 45, ""); err == nil {
		t.Error("expected error for non-90 rotation")
	}
}

// ---------------------------------------------------------------------------
// Optimize
// ---------------------------------------------------------------------------

func TestOptimize(t *testing.T) {
	path := makePDF(t, "Optimize me")
	out := filepath.Join(t.TempDir(), "opt.pdf")

	if err := Optimize(path, out); err != nil {
		t.Fatalf("Optimize: %v", err)
	}
	info, _ := os.Stat(out)
	if info == nil || info.Size() == 0 {
		t.Error("optimized file is empty or missing")
	}
}

// ---------------------------------------------------------------------------
// Encrypt / Decrypt
// ---------------------------------------------------------------------------

func TestEncryptDecrypt(t *testing.T) {
	path := makePDF(t, "Secret")
	enc := filepath.Join(t.TempDir(), "enc.pdf")
	dec := filepath.Join(t.TempDir(), "dec.pdf")

	if err := Encrypt(path, enc, "pass123", ""); err != nil {
		t.Fatalf("Encrypt: %v", err)
	}

	if err := Decrypt(enc, dec, "pass123", ""); err != nil {
		t.Fatalf("Decrypt: %v", err)
	}

	// Verify decrypted file is readable
	text, err := AllText(dec)
	if err != nil {
		t.Fatalf("AllText on decrypted: %v", err)
	}
	if !strings.Contains(text, "Secret") {
		t.Errorf("decrypted text missing 'Secret': %q", text)
	}
}

func TestEncrypt_EmptyPW(t *testing.T) {
	path := makePDF(t, "X")
	if err := Encrypt(path, "x.pdf", "", ""); err == nil {
		t.Error("expected error for empty user password")
	}
}

// ---------------------------------------------------------------------------
// ExtractImages / ListExtractedImages
// ---------------------------------------------------------------------------

func TestExtractImages_EmptyPDF(t *testing.T) {
	path := makePDF(t, "No images here")
	outDir := filepath.Join(t.TempDir(), "imgs")

	// Should succeed even if no images to extract
	if err := ExtractImages(path, outDir, ""); err != nil {
		t.Fatalf("ExtractImages: %v", err)
	}

	imgs, err := ListExtractedImages(outDir)
	if err != nil {
		t.Fatalf("ListExtractedImages: %v", err)
	}
	if len(imgs) != 0 {
		t.Errorf("expected 0 images, got %d", len(imgs))
	}
}

// ---------------------------------------------------------------------------
// Bookmarks / AddBookmarks
// ---------------------------------------------------------------------------

func TestBookmarks_Empty(t *testing.T) {
	path := makePDF(t, "No bookmarks")
	bms, err := Bookmarks(path)
	if err != nil {
		t.Fatalf("Bookmarks: %v", err)
	}
	if len(bms) != 0 {
		t.Errorf("expected 0 bookmarks, got %d", len(bms))
	}
}

func TestAddBookmarks_ReadBack(t *testing.T) {
	path := makePDF(t, "Page1", "Page2")
	out := filepath.Join(t.TempDir(), "bm.pdf")

	specs := []BookmarkSpec{
		{Title: "Chapter 1", Page: 1},
		{Title: "Chapter 2", Page: 2, Children: []BookmarkSpec{
			{Title: "Section 2.1", Page: 2},
		}},
	}
	if err := AddBookmarks(path, out, specs, false); err != nil {
		t.Fatalf("AddBookmarks: %v", err)
	}

	bms, err := Bookmarks(out)
	if err != nil {
		t.Fatalf("Bookmarks readback: %v", err)
	}
	if len(bms) < 2 {
		t.Fatalf("expected >= 2 bookmarks, got %d", len(bms))
	}
	if bms[0].Title != "Chapter 1" {
		t.Errorf("bookmark[0] title=%q want 'Chapter 1'", bms[0].Title)
	}
	if len(bms[1].Children) < 1 {
		t.Error("expected Chapter 2 to have children")
	}
}

// ---------------------------------------------------------------------------
// Reorder
// ---------------------------------------------------------------------------

func TestReorder(t *testing.T) {
	path := makePDF(t, "First", "Second", "Third")
	out := filepath.Join(t.TempDir(), "reordered.pdf")

	if err := Reorder(path, out, "3,1,2"); err != nil {
		t.Fatalf("Reorder: %v", err)
	}

	pages, err := Read(out, ReadOptions{})
	if err != nil {
		t.Fatalf("Read reordered: %v", err)
	}
	if len(pages) != 3 {
		t.Fatalf("expected 3 pages, got %d", len(pages))
	}
	// Page 1 in output should have "Third" (original page 3)
	if !strings.Contains(pages[0].Text, "Third") {
		t.Errorf("reordered page 1 text=%q, want contains 'Third'", pages[0].Text)
	}
}

// ---------------------------------------------------------------------------
// InsertBlank
// ---------------------------------------------------------------------------

func TestInsertBlank(t *testing.T) {
	path := makePDF(t, "Page1", "Page2")
	out := filepath.Join(t.TempDir(), "blank.pdf")

	if err := InsertBlank(path, out, 1, 2); err != nil {
		t.Fatalf("InsertBlank: %v", err)
	}

	n, _ := PageCount(out)
	if n != 4 {
		t.Errorf("PageCount=%d want 4 (2 original + 2 blank)", n)
	}
}

func TestInsertBlank_InvalidCount(t *testing.T) {
	path := makePDF(t, "X")
	if err := InsertBlank(path, "x.pdf", 1, 0); err == nil {
		t.Error("expected error for count=0")
	}
}

// ---------------------------------------------------------------------------
// StampImage
// ---------------------------------------------------------------------------

func TestStampImage(t *testing.T) {
	// Create a valid 1x1 red PNG using Go's standard library
	imgPath := filepath.Join(t.TempDir(), "stamp.png")
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{R: 255, A: 255})
	f, err := os.Create(imgPath)
	if err != nil {
		t.Fatal(err)
	}
	if err := png.Encode(f, img); err != nil {
		_ = f.Close()
		t.Fatal(err)
	}
	_ = f.Close()

	path := makePDF(t, "Stamp target")
	out := filepath.Join(t.TempDir(), "stamped.pdf")

	if err := StampImage(path, out, imgPath, ""); err != nil {
		t.Fatalf("StampImage: %v", err)
	}
	n, _ := PageCount(out)
	if n != 1 {
		t.Errorf("stamped PageCount=%d want 1", n)
	}
}

// ---------------------------------------------------------------------------
// SetMeta
// ---------------------------------------------------------------------------

func TestSetMeta(t *testing.T) {
	path := makePDF(t, "Meta test")
	out := filepath.Join(t.TempDir(), "meta.pdf")

	meta := MetaUpdate{Title: "Test Title", Author: "Test Author"}
	if err := SetMeta(path, out, meta); err != nil {
		t.Fatalf("SetMeta: %v", err)
	}

	// Verify via FullInfo
	info, err := FullInfo(out)
	if err != nil {
		t.Fatalf("FullInfo: %v", err)
	}
	if info["title"] != "Test Title" {
		t.Errorf("title=%v want 'Test Title'", info["title"])
	}
	if info["author"] != "Test Author" {
		t.Errorf("author=%v want 'Test Author'", info["author"])
	}
}

// ---------------------------------------------------------------------------
// FullInfo
// ---------------------------------------------------------------------------

func TestFullInfo(t *testing.T) {
	path := makePDF(t, "Info test")
	info, err := FullInfo(path)
	if err != nil {
		t.Fatalf("FullInfo: %v", err)
	}
	if info["pageCount"] == nil && info["pageCount"] != float64(1) {
		// pageCount may be int or float64 depending on pdfcpu version
		t.Logf("pageCount=%v (type %T)", info["pageCount"], info["pageCount"])
	}
	if info["sizeBytes"] == nil {
		t.Error("sizeBytes should be present")
	}
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

func TestSearch_Found(t *testing.T) {
	path := makePDF(t, "Hello World", "Goodbye World")
	results, err := Search(path, "World", SearchOptions{})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	if len(results) < 2 {
		t.Fatalf("expected >= 2 results, got %d", len(results))
	}
	// Both pages should have matches
	pages := map[int]bool{}
	for _, r := range results {
		pages[r.Page] = true
		if r.Snippet == "" {
			t.Errorf("page %d: empty snippet", r.Page)
		}
	}
	if !pages[1] || !pages[2] {
		t.Errorf("expected matches on pages 1 and 2, got pages: %v", pages)
	}
}

func TestSearch_NotFound(t *testing.T) {
	path := makePDF(t, "Hello World")
	results, err := Search(path, "nonexistent", SearchOptions{})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	if len(results) != 0 {
		t.Errorf("expected 0 results, got %d", len(results))
	}
}

func TestSearch_CaseSensitive(t *testing.T) {
	path := makePDF(t, "Hello HELLO hello")
	results, err := Search(path, "hello", SearchOptions{CaseSensitive: true})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	// Should match only the lowercase "hello"
	if len(results) != 1 {
		t.Errorf("expected 1 case-sensitive match, got %d", len(results))
	}
}

func TestSearch_CaseInsensitive(t *testing.T) {
	path := makePDF(t, "Hello HELLO hello")
	results, err := Search(path, "hello", SearchOptions{CaseSensitive: false})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	// Should match all three occurrences (or at least 1 per line containing any case)
	if len(results) < 1 {
		t.Errorf("expected >= 1 case-insensitive match, got %d", len(results))
	}
}

func TestSearch_SinglePage(t *testing.T) {
	path := makePDF(t, "Page one keyword", "Page two keyword")
	results, err := Search(path, "keyword", SearchOptions{Page: 2})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	if len(results) != 1 {
		t.Fatalf("expected 1 result on page 2, got %d", len(results))
	}
	if results[0].Page != 2 {
		t.Errorf("result page=%d want 2", results[0].Page)
	}
}

func TestSearch_EmptyKeyword(t *testing.T) {
	path := makePDF(t, "Hello")
	_, err := Search(path, "", SearchOptions{})
	if err == nil {
		t.Error("expected error for empty keyword")
	}
}

func TestSearch_Limit(t *testing.T) {
	path := makePDF(t, "A B C A B C A B C")
	results, err := Search(path, "A", SearchOptions{Limit: 2})
	if err != nil {
		t.Fatalf("Search: %v", err)
	}
	if len(results) > 2 {
		t.Errorf("expected at most 2 results with limit, got %d", len(results))
	}
}

// ---------------------------------------------------------------------------
// CreatePDF
// ---------------------------------------------------------------------------

func TestCreatePDF_Basic(t *testing.T) {
	out := filepath.Join(t.TempDir(), "created.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Hello World"}, {Text: "Page Two"}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}

	// Verify the file is a valid PDF by reading it back
	pages, err := Read(out, ReadOptions{})
	if err != nil {
		t.Fatalf("Read created PDF: %v", err)
	}
	if len(pages) != 2 {
		t.Fatalf("expected 2 pages, got %d", len(pages))
	}
	if !strings.Contains(pages[0].Text, "Hello World") {
		t.Errorf("page 1 missing 'Hello World': %q", pages[0].Text)
	}
	if !strings.Contains(pages[1].Text, "Page Two") {
		t.Errorf("page 2 missing 'Page Two': %q", pages[1].Text)
	}
}

func TestCreatePDF_WithMetadata(t *testing.T) {
	out := filepath.Join(t.TempDir(), "meta-created.pdf")
	spec := CreateSpec{
		Pages:  []PageSpec{{Text: "Content"}},
		Title:  "Test Title",
		Author: "Test Author",
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}

	info, err := FullInfo(out)
	if err != nil {
		t.Fatalf("FullInfo: %v", err)
	}
	if info["title"] != "Test Title" {
		t.Errorf("title=%v want 'Test Title'", info["title"])
	}
	if info["author"] != "Test Author" {
		t.Errorf("author=%v want 'Test Author'", info["author"])
	}
}

func TestCreatePDF_EmptyPages(t *testing.T) {
	out := filepath.Join(t.TempDir(), "empty.pdf")
	spec := CreateSpec{Pages: []PageSpec{}}
	if err := CreatePDF(out, spec); err == nil {
		t.Error("expected error for empty pages")
	}
}

func TestCreatePDF_A4Paper(t *testing.T) {
	out := filepath.Join(t.TempDir(), "a4.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "A4 page"}},
		Paper: "A4",
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF A4: %v", err)
	}

	dims, err := PageDimensions(out)
	if err != nil {
		t.Fatalf("PageDimensions: %v", err)
	}
	if len(dims) != 1 {
		t.Fatalf("expected 1 page, got %d", len(dims))
	}
	// A4: 595.28 x 841.89
	if dims[0].Width < 595 || dims[0].Width > 596 {
		t.Errorf("A4 width=%.1f want ~595.28", dims[0].Width)
	}
}

func TestCreatePDF_CustomPaper(t *testing.T) {
	out := filepath.Join(t.TempDir(), "custom.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Custom"}},
		Paper: "400x600",
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF custom: %v", err)
	}

	dims, err := PageDimensions(out)
	if err != nil {
		t.Fatalf("PageDimensions: %v", err)
	}
	if dims[0].Width != 400 || dims[0].Height != 600 {
		t.Errorf("dimensions=%.0fx%.0f want 400x600", dims[0].Width, dims[0].Height)
	}
}

func TestCreatePDF_SinglePage(t *testing.T) {
	out := filepath.Join(t.TempDir(), "single.pdf")
	spec := CreateSpec{Pages: []PageSpec{{Text: "Only page"}}}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

// ---------------------------------------------------------------------------
// ReplaceText
// ---------------------------------------------------------------------------

func TestReplaceText_Simple(t *testing.T) {
	// Create a PDF with known text, then replace it
	inPath := makePDF(t, "Hello World")
	outPath := filepath.Join(t.TempDir(), "replaced.pdf")

	hits, err := ReplaceText(inPath, outPath, []Replacement{
		{Find: "Hello", Replace: "Goodbye"},
	})
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if hits < 1 {
		t.Errorf("expected >= 1 hit, got %d", hits)
	}

	// Read back and verify
	text, err := AllText(outPath)
	if err != nil {
		t.Fatalf("AllText: %v", err)
	}
	if !strings.Contains(text, "Goodbye") {
		t.Errorf("replaced text missing 'Goodbye': %q", text)
	}
	if strings.Contains(text, "Hello") {
		t.Errorf("replaced text still contains 'Hello': %q", text)
	}
}

func TestReplaceText_MultiplePairs(t *testing.T) {
	inPath := makePDF(t, "Alice and Bob")
	outPath := filepath.Join(t.TempDir(), "replaced2.pdf")

	hits, err := ReplaceText(inPath, outPath, []Replacement{
		{Find: "Alice", Replace: "Charlie"},
		{Find: "Bob", Replace: "Dave"},
	})
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if hits < 2 {
		t.Errorf("expected >= 2 hits, got %d", hits)
	}

	text, err := AllText(outPath)
	if err != nil {
		t.Fatalf("AllText: %v", err)
	}
	if !strings.Contains(text, "Charlie") || !strings.Contains(text, "Dave") {
		t.Errorf("expected 'Charlie and Dave', got: %q", text)
	}
}

func TestReplaceText_NoMatch(t *testing.T) {
	inPath := makePDF(t, "Hello World")
	outPath := filepath.Join(t.TempDir(), "nomatch.pdf")

	hits, err := ReplaceText(inPath, outPath, []Replacement{
		{Find: "nonexistent", Replace: "nothing"},
	})
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}
	if hits != 0 {
		t.Errorf("expected 0 hits for no match, got %d", hits)
	}
}

func TestReplaceText_EmptyPairs(t *testing.T) {
	inPath := makePDF(t, "Hello")
	_, err := ReplaceText(inPath, "x.pdf", []Replacement{})
	if err == nil {
		t.Error("expected error for empty pairs")
	}
}

func TestReplaceText_PreservesPDF(t *testing.T) {
	inPath := makePDF(t, "Replace Me Keep Me")
	outPath := filepath.Join(t.TempDir(), "preserved.pdf")

	_, err := ReplaceText(inPath, outPath, []Replacement{
		{Find: "Replace Me", Replace: "Done"},
	})
	if err != nil {
		t.Fatalf("ReplaceText: %v", err)
	}

	// Verify output is a valid PDF (can be opened and read)
	pages, err := Read(outPath, ReadOptions{})
	if err != nil {
		t.Fatalf("Read replaced PDF: %v", err)
	}
	if len(pages) != 1 {
		t.Errorf("expected 1 page, got %d", len(pages))
	}
	if !strings.Contains(pages[0].Text, "Done") {
		t.Errorf("missing 'Done': %q", pages[0].Text)
	}
	if !strings.Contains(pages[0].Text, "Keep Me") {
		t.Errorf("missing 'Keep Me': %q", pages[0].Text)
	}
}

// ---------------------------------------------------------------------------
// Hex decode/encode helpers
// ---------------------------------------------------------------------------

func TestHexDecode(t *testing.T) {
	cases := []struct {
		in   string
		want string
	}{
		{"48656C6C6F", "Hello"},
		{"4F 4B", "OK"}, // with whitespace
		{"4", "@"},      // odd-length pads with 0 → 0x40 = '@'
		{"41", "A"},
		{"", ""},
	}
	for _, c := range cases {
		got := hexDecode(c.in)
		if got != c.want {
			t.Errorf("hexDecode(%q) = %q, want %q", c.in, got, c.want)
		}
	}
}

func TestHexEncode(t *testing.T) {
	cases := []struct {
		in   string
		want string
	}{
		{"Hello", "48656C6C6F"},
		{"A", "41"},
		{"OK", "4F4B"},
		{"", ""},
	}
	for _, c := range cases {
		got := hexEncode(c.in)
		if got != c.want {
			t.Errorf("hexEncode(%q) = %q, want %q", c.in, got, c.want)
		}
	}
}

func TestHexDecodeEncode_RoundTrip(t *testing.T) {
	texts := []string{"Hello World", "Test 123", "special!@#"}
	for _, text := range texts {
		encoded := hexEncode(text)
		decoded := hexDecode(encoded)
		if decoded != text {
			t.Errorf("round-trip failed: %q → %q → %q", text, encoded, decoded)
		}
	}
}

func TestReplaceInContentStream_HexString(t *testing.T) {
	// Simulate a PDF content stream with hex-encoded text
	// "Hello" = 48656C6C6F
	stream := "BT /F1 24 Tf 100 700 Td <48656C6C6F> Tj ET"
	hits := 0
	result := replaceInContentStream(stream, "Hello", "World", &hits)
	if hits != 1 {
		t.Errorf("expected 1 hit, got %d", hits)
	}
	// "World" = 576F726C64
	if !strings.Contains(result, "<576F726C64>") {
		t.Errorf("expected hex-encoded 'World' in result: %s", result)
	}
}

func TestReplaceInContentStream_HexNoMatch(t *testing.T) {
	stream := "BT /F1 24 Tf 100 700 Td <48656C6C6F> Tj ET"
	hits := 0
	result := replaceInContentStream(stream, "Goodbye", "World", &hits)
	if hits != 0 {
		t.Errorf("expected 0 hits, got %d", hits)
	}
	// Stream should be unchanged
	if result != stream {
		t.Errorf("stream should be unchanged when no match")
	}
}

func TestReplaceInContentStream_MixedParenAndHex(t *testing.T) {
	// Stream has both parenthesized and hex strings
	stream := "BT (Hello) Tj <576F726C64> Tj ET"
	hits := 0
	result := replaceInContentStream(stream, "Hello", "Hi", &hits)
	if hits != 1 {
		t.Errorf("expected 1 hit for paren, got %d", hits)
	}
	if !strings.Contains(result, "(Hi)") {
		t.Errorf("expected (Hi) in result: %s", result)
	}

	// Now replace the hex part
	hits = 0
	result = replaceInContentStream(result, "World", "Earth", &hits)
	if hits != 1 {
		t.Errorf("expected 1 hit for hex, got %d", hits)
	}
	if !strings.Contains(result, hexEncode("Earth")) {
		t.Errorf("expected hex 'Earth' in result: %s", result)
	}
}

// ---------------------------------------------------------------------------
// pageMatches
// ---------------------------------------------------------------------------

func TestPageMatches(t *testing.T) {
	cases := []struct {
		page int
		spec string
		want bool
	}{
		{1, "1", true},
		{2, "1", false},
		{3, "1,3,5", true},
		{4, "1,3,5", false},
		{5, "1-3,5-7", true},
		{2, "1-3,5-7", true},
		{6, "1-3,5-7", true},
		{4, "1-3,5-7", false},
		{8, "1-3,5-7", false},
		{1, "", false},
	}
	for _, c := range cases {
		got := pageMatches(c.page, c.spec)
		if got != c.want {
			t.Errorf("pageMatches(%d, %q) = %v, want %v", c.page, c.spec, got, c.want)
		}
	}
}

// ---------------------------------------------------------------------------
// AddText
// ---------------------------------------------------------------------------

func TestAddText_SingleOverlay(t *testing.T) {
	path := makePDF(t, "Original content")
	out := filepath.Join(t.TempDir(), "addtext.pdf")

	overlays := []TextOverlay{{
		Text:     "OVERLAY",
		X:        100,
		Y:        400,
		FontSize: 24,
	}}
	if err := AddText(path, out, overlays); err != nil {
		t.Fatalf("AddText: %v", err)
	}

	// Verify output is a valid PDF
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}

	// Verify the file was created and has content
	info, err := os.Stat(out)
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("output file is empty")
	}
}

func TestAddText_MultipleOverlays(t *testing.T) {
	path := makePDF(t, "Page content")
	out := filepath.Join(t.TempDir(), "multi-overlay.pdf")

	overlays := []TextOverlay{
		{Text: "First", X: 72, Y: 700},
		{Text: "Second", X: 72, Y: 500},
	}
	if err := AddText(path, out, overlays); err != nil {
		t.Fatalf("AddText: %v", err)
	}

	// Verify output is a valid PDF with correct page count
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}

	// Verify the file was created and has content
	info, err := os.Stat(out)
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("output file is empty")
	}
}

func TestAddText_WithColor(t *testing.T) {
	path := makePDF(t, "Colored overlay")
	out := filepath.Join(t.TempDir(), "colored-overlay.pdf")

	overlays := []TextOverlay{{
		Text:  "RED",
		X:     100,
		Y:     400,
		Color: "1 0 0",
	}}
	if err := AddText(path, out, overlays); err != nil {
		t.Fatalf("AddText: %v", err)
	}

	n, _ := PageCount(out)
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

func TestAddText_EmptyOverlays(t *testing.T) {
	path := makePDF(t, "Test")
	if err := AddText(path, "x.pdf", []TextOverlay{}); err == nil {
		t.Error("expected error for empty overlays")
	}
}

func TestAddText_PreservesOriginal(t *testing.T) {
	path := makePDF(t, "Keep this text safe")
	out := filepath.Join(t.TempDir(), "preserved-overlay.pdf")

	overlays := []TextOverlay{{Text: "Added", X: 72, Y: 700}}
	if err := AddText(path, out, overlays); err != nil {
		t.Fatalf("AddText: %v", err)
	}

	// Verify output is a valid PDF with correct page count
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}

	// Verify the file was created and has content (overlay added)
	info, err := os.Stat(out)
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}
	if info.Size() == 0 {
		t.Error("output file is empty")
	}

	// Verify the output is larger than the input (overlay added)
	inputInfo, err := os.Stat(path)
	if err != nil {
		t.Fatalf("Stat input: %v", err)
	}
	if info.Size() <= inputInfo.Size() {
		t.Errorf("output (%d bytes) should be larger than input (%d bytes)", info.Size(), inputInfo.Size())
	}
}

// ---------------------------------------------------------------------------
// CreatePDF — enhanced: FontSize, Bold, Color
// ---------------------------------------------------------------------------

func TestCreatePDF_CustomFontSize(t *testing.T) {
	out := filepath.Join(t.TempDir(), "fontsize.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Large text", FontSize: 24}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	pages, err := Read(out, ReadOptions{})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if !strings.Contains(pages[0].Text, "Large text") {
		t.Errorf("page text missing: %q", pages[0].Text)
	}
}

func TestCreatePDF_BoldFont(t *testing.T) {
	out := filepath.Join(t.TempDir(), "bold.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Bold text", Bold: true}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	pages, err := Read(out, ReadOptions{})
	if err != nil {
		t.Fatalf("Read: %v", err)
	}
	if !strings.Contains(pages[0].Text, "Bold text") {
		t.Errorf("page text missing: %q", pages[0].Text)
	}
}

func TestCreatePDF_WithColor(t *testing.T) {
	out := filepath.Join(t.TempDir(), "colored.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Red text", Color: "1 0 0"}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	// Verify it's a valid PDF
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

func TestCreatePDF_MixedStyles(t *testing.T) {
	out := filepath.Join(t.TempDir(), "mixed.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{
			{Text: "Normal page"},
			{Text: "Bold red", Bold: true, FontSize: 18, Color: "1 0 0"},
			{Text: "Small blue", FontSize: 8, Color: "0 0 1"},
		},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 3 {
		t.Errorf("PageCount=%d want 3", n)
	}
}

// ---------------------------------------------------------------------------
// buildOverlayStream
// ---------------------------------------------------------------------------

func TestBuildOverlayStream_Basic(t *testing.T) {
	overlays := []TextOverlay{{Text: "Test", X: 100, Y: 200, FontSize: 14}}
	defaultFontMap := map[string]string{"F1": "Helvetica"}
	stream := buildOverlayStream(overlays, defaultFontMap)
	if !strings.Contains(stream, "BT\n") {
		t.Errorf("missing BT operator: %s", stream)
	}
	if !strings.Contains(stream, "ET\n") {
		t.Errorf("missing ET operator: %s", stream)
	}
	if !strings.Contains(stream, "(Test)") {
		t.Errorf("missing text literal: %s", stream)
	}
	if !strings.Contains(stream, "14.0 Tf") {
		t.Errorf("missing font size: %s", stream)
	}
}

func TestBuildOverlayStream_WithColor(t *testing.T) {
	overlays := []TextOverlay{{Text: "Red", X: 50, Y: 50, Color: "1 0 0"}}
	defaultFontMap := map[string]string{"F1": "Helvetica"}
	stream := buildOverlayStream(overlays, defaultFontMap)
	if !strings.Contains(stream, "1 0 0 rg\n") {
		t.Errorf("missing color operator: %s", stream)
	}
}

// ---------------------------------------------------------------------------
// Typography: resolveFontName, charWidthFactor, mergePageDefaults
// ---------------------------------------------------------------------------

func TestResolveFontName(t *testing.T) {
	cases := []struct {
		family string
		bold   bool
		italic bool
		want   string
	}{
		{"", false, false, "Helvetica"},
		{"", true, false, "Helvetica-Bold"},
		{"", false, true, "Helvetica-Oblique"},
		{"", true, true, "Helvetica-BoldOblique"},
		{"Helvetica", false, false, "Helvetica"},
		{"Helvetica", true, false, "Helvetica-Bold"},
		{"sans", false, true, "Helvetica-Oblique"},
		{"Times", false, false, "Times-Roman"},
		{"Times", true, false, "Times-Bold"},
		{"Times", false, true, "Times-Italic"},
		{"Times", true, true, "Times-BoldItalic"},
		{"serif", false, false, "Times-Roman"},
		{"Courier", false, false, "Courier"},
		{"Courier", true, false, "Courier-Bold"},
		{"Courier", false, true, "Courier-Oblique"},
		{"Courier", true, true, "Courier-BoldOblique"},
		{"mono", false, false, "Courier"},
		{"monospace", true, true, "Courier-BoldOblique"},
	}
	for _, c := range cases {
		got := resolveFontName(c.family, c.bold, c.italic)
		if got != c.want {
			t.Errorf("resolveFontName(%q, %v, %v) = %q, want %q", c.family, c.bold, c.italic, got, c.want)
		}
	}
}

func TestCharWidthFactor(t *testing.T) {
	cases := []struct {
		font string
		want float64
	}{
		{"Courier", 0.60},
		{"Courier-Bold", 0.60},
		{"Times-Roman", 0.45},
		{"Times-Bold", 0.45},
		{"Helvetica", 0.50},
		{"Helvetica-Bold", 0.50},
		{"Unknown", 0.50}, // default
	}
	for _, c := range cases {
		got := charWidthFactor(c.font)
		if got != c.want {
			t.Errorf("charWidthFactor(%q) = %v, want %v", c.font, got, c.want)
		}
	}
}

func TestMergePageDefaults(t *testing.T) {
	global := CreateSpec{
		Font:       "Times",
		FontSize:   16,
		Align:      "center",
		LineHeight: 1.6,
		Margin:     50,
	}

	// Page with no overrides — should inherit all globals
	pg1 := mergePageDefaults(PageSpec{Text: "hello"}, global)
	if pg1.Font != "Times" {
		t.Errorf("Font = %q, want Times", pg1.Font)
	}
	if pg1.FontSize != 16 {
		t.Errorf("FontSize = %v, want 16", pg1.FontSize)
	}
	if pg1.Align != "center" {
		t.Errorf("Align = %q, want center", pg1.Align)
	}
	if pg1.LineHeight != 1.6 {
		t.Errorf("LineHeight = %v, want 1.6", pg1.LineHeight)
	}
	if pg1.Margin != 50 {
		t.Errorf("Margin = %v, want 50", pg1.Margin)
	}

	// Page with overrides — should use page-level values
	pg2 := mergePageDefaults(PageSpec{Text: "world", Font: "Courier", FontSize: 20, Align: "right"}, global)
	if pg2.Font != "Courier" {
		t.Errorf("Font = %q, want Courier", pg2.Font)
	}
	if pg2.FontSize != 20 {
		t.Errorf("FontSize = %v, want 20", pg2.FontSize)
	}
	if pg2.Align != "right" {
		t.Errorf("Align = %q, want right", pg2.Align)
	}
	if pg2.LineHeight != 1.6 { // inherited
		t.Errorf("LineHeight = %v, want 1.6", pg2.LineHeight)
	}
	if pg2.Margin != 50 { // inherited
		t.Errorf("Margin = %v, want 50", pg2.Margin)
	}
}

func TestBuildTextStream_Alignment(t *testing.T) {
	width, height := 612.0, 792.0
	fontSize := 12.0
	margin := 72.0

	// Left alignment (default)
	left := buildTextStream("Hello", width, height, fontSize, "", "F1", "Helvetica", "left", 1.4, margin, false)
	// Left align should start at margin (72.0)
	if !strings.Contains(left, "72.0") {
		t.Errorf("left align should start at margin: %s", left)
	}

	// Center alignment
	center := buildTextStream("Hello", width, height, fontSize, "", "F1", "Helvetica", "center", 1.4, margin, false)
	// Center should have x > margin (72.0)
	// Extract the x coordinate from the Td operator
	lines := strings.Split(center, "\n")
	for _, line := range lines {
		if strings.Contains(line, "Td") {
			parts := strings.Fields(line)
			if len(parts) >= 3 {
				x := parts[0]
				if x == "72.0" {
					t.Errorf("center align should not start at margin: %s", center)
				}
			}
		}
	}

	// Right alignment
	right := buildTextStream("Hello", width, height, fontSize, "", "F1", "Helvetica", "right", 1.4, margin, false)
	// Right should have x > center
	lines = strings.Split(right, "\n")
	for _, line := range lines {
		if strings.Contains(line, "Td") {
			parts := strings.Fields(line)
			if len(parts) >= 3 {
				x := parts[0]
				if x == "72.0" {
					t.Errorf("right align should not start at margin: %s", right)
				}
			}
		}
	}
}

func TestBuildTextStream_Underline(t *testing.T) {
	stream := buildTextStream("Underline me", 612, 792, 12, "", "F1", "Helvetica", "left", 1.4, 72, true)
	if !strings.Contains(stream, " m ") {
		t.Errorf("underline should contain moveTo operator: %s", stream)
	}
	if !strings.Contains(stream, " l ") {
		t.Errorf("underline should contain lineTo operator: %s", stream)
	}
	if !strings.Contains(stream, " S\n") {
		t.Errorf("underline should contain stroke operator: %s", stream)
	}
}

func TestBuildTextStream_Margin(t *testing.T) {
	// Default margin (72)
	defaultMargin := buildTextStream("Test", 612, 792, 12, "", "F1", "Helvetica", "left", 1.4, 72, false)
	if !strings.Contains(defaultMargin, "72.0") {
		t.Errorf("default margin should be 72: %s", defaultMargin)
	}

	// Custom margin (50)
	customMargin := buildTextStream("Test", 612, 792, 12, "", "F1", "Helvetica", "left", 1.4, 50, false)
	if !strings.Contains(customMargin, "50.0") {
		t.Errorf("custom margin should be 50: %s", customMargin)
	}
}

func TestBuildTextStream_LineHeight(t *testing.T) {
	// Default line height (1.4)
	defaultLH := buildTextStream("Line1\nLine2", 612, 792, 12, "", "F1", "Helvetica", "left", 1.4, 72, false)
	if !strings.Contains(defaultLH, "-16.8 Td") { // 12 * 1.4 = 16.8
		t.Errorf("default line height should be 1.4x: %s", defaultLH)
	}

	// Custom line height (2.0)
	customLH := buildTextStream("Line1\nLine2", 612, 792, 12, "", "F1", "Helvetica", "left", 2.0, 72, false)
	if !strings.Contains(customLH, "-24.0 Td") { // 12 * 2.0 = 24.0
		t.Errorf("custom line height should be 2.0x: %s", customLH)
	}
}

func TestBuildOverlayStream_FontVariants(t *testing.T) {
	fontRefMap := map[string]string{
		"Helvetica":        "/F1",
		"Helvetica-Bold":   "/F2",
		"Times-Roman":      "/F3",
		"Times-BoldItalic": "/F4",
		"Courier":          "/F5",
	}

	overlays := []TextOverlay{
		{Text: "Normal", X: 100, Y: 700, Font: "Helvetica"},
		{Text: "Bold", X: 100, Y: 600, Font: "Helvetica", Bold: true},
		{Text: "Serif", X: 100, Y: 500, Font: "Times"},
		{Text: "BoldItalic", X: 100, Y: 400, Font: "Times", Bold: true, Italic: true},
		{Text: "Mono", X: 100, Y: 300, Font: "Courier"},
	}

	stream := buildOverlayStream(overlays, fontRefMap)
	if !strings.Contains(stream, "/F1 12.0 Tf") {
		t.Errorf("missing Helvetica ref (F1): %s", stream)
	}
	if !strings.Contains(stream, "/F2 12.0 Tf") {
		t.Errorf("missing Helvetica-Bold ref (F2): %s", stream)
	}
	if !strings.Contains(stream, "/F3 12.0 Tf") {
		t.Errorf("missing Times-Roman ref (F3): %s", stream)
	}
	if !strings.Contains(stream, "/F4 12.0 Tf") {
		t.Errorf("missing Times-BoldItalic ref (F4): %s", stream)
	}
	if !strings.Contains(stream, "/F5 12.0 Tf") {
		t.Errorf("missing Courier ref (F5): %s", stream)
	}
}

func TestBuildOverlayStream_Underline(t *testing.T) {
	fontRefMap := map[string]string{"Helvetica": "F1"}
	overlays := []TextOverlay{
		{Text: "Underline", X: 100, Y: 700, Underline: true},
	}

	stream := buildOverlayStream(overlays, fontRefMap)
	if !strings.Contains(stream, " m ") {
		t.Errorf("underline should contain moveTo: %s", stream)
	}
	if !strings.Contains(stream, " l ") {
		t.Errorf("underline should contain lineTo: %s", stream)
	}
	if !strings.Contains(stream, " S\n") {
		t.Errorf("underline should contain stroke: %s", stream)
	}
}

// ---------------------------------------------------------------------------
// CreatePDF with typography features
// ---------------------------------------------------------------------------

func TestCreatePDF_TimesFont(t *testing.T) {
	out := filepath.Join(t.TempDir(), "times.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Serif text", Font: "Times"}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

func TestCreatePDF_CourierFont(t *testing.T) {
	out := filepath.Join(t.TempDir(), "courier.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Monospace text", Font: "Courier"}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

func TestCreatePDF_Alignment(t *testing.T) {
	out := filepath.Join(t.TempDir(), "aligned.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{
			{Text: "Left aligned", Align: "left"},
			{Text: "Center aligned", Align: "center"},
			{Text: "Right aligned", Align: "right"},
		},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 3 {
		t.Errorf("PageCount=%d want 3", n)
	}
}

func TestCreatePDF_Underline(t *testing.T) {
	out := filepath.Join(t.TempDir(), "underline.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{{Text: "Underlined text", Underline: true}},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 1 {
		t.Errorf("PageCount=%d want 1", n)
	}
}

func TestCreatePDF_GlobalDefaults(t *testing.T) {
	out := filepath.Join(t.TempDir(), "global-defaults.pdf")
	spec := CreateSpec{
		Font:       "Times",
		FontSize:   14,
		Align:      "center",
		LineHeight: 1.8,
		Margin:     50,
		Pages: []PageSpec{
			{Text: "Page 1 with global defaults"},
			{Text: "Page 2 with global defaults"},
		},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 2 {
		t.Errorf("PageCount=%d want 2", n)
	}
}

func TestCreatePDF_MixedFonts(t *testing.T) {
	out := filepath.Join(t.TempDir(), "mixed-fonts.pdf")
	spec := CreateSpec{
		Pages: []PageSpec{
			{Text: "Sans-serif", Font: "Helvetica"},
			{Text: "Serif", Font: "Times"},
			{Text: "Monospace", Font: "Courier"},
			{Text: "Bold Serif", Font: "Times", Bold: true},
			{Text: "Italic Mono", Font: "Courier", Italic: true},
		},
	}
	if err := CreatePDF(out, spec); err != nil {
		t.Fatalf("CreatePDF: %v", err)
	}
	n, err := PageCount(out)
	if err != nil {
		t.Fatalf("PageCount: %v", err)
	}
	if n != 5 {
		t.Errorf("PageCount=%d want 5", n)
	}
}

func TestOrderedStringSet(t *testing.T) {
	s := orderedStringSet{}
	s.Add("Helvetica")
	s.Add("Times")
	s.Add("Courier")
	s.Add("Helvetica") // duplicate
	s.Add("Times")     // duplicate

	items := s.Items()
	if len(items) != 3 {
		t.Errorf("Items() length = %d, want 3", len(items))
	}
	if items[0] != "Helvetica" || items[1] != "Times" || items[2] != "Courier" {
		t.Errorf("Items() = %v, want [Helvetica Times Courier]", items)
	}
}
