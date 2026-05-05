// Package pdf provides PDF read / merge / split / trim / watermark operations.
//
// Reading text content uses github.com/ledongthuc/pdf (lightweight, content-stream
// based extractor). Structural operations (merge, split, watermark, info) use
// github.com/pdfcpu/pdfcpu, which is the most feature-complete pure-Go PDF library.
package pdf

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	ledongthucpdf "github.com/ledongthuc/pdf"
	pdfapi "github.com/pdfcpu/pdfcpu/pkg/api"
	pdfcpu "github.com/pdfcpu/pdfcpu/pkg/pdfcpu"
	pdfmodel "github.com/pdfcpu/pdfcpu/pkg/pdfcpu/model"
	pdftypes "github.com/pdfcpu/pdfcpu/pkg/pdfcpu/types"
)

// Page is one page's extracted text and metadata.
type Page struct {
	Page      int    `json:"page"`
	Text      string `json:"text"`
	WordCount int    `json:"wordCount"`
}

// ReadOptions narrow Read to a specific page or range.
type ReadOptions struct {
	From  int // 1-based; 0 = start at first
	To    int // 0 = until last
	Limit int // max pages to return; 0 = no limit
}

// Read returns the text of (each page in) the PDF file, optionally limited by range.
func Read(path string, opts ReadOptions) ([]Page, error) {
	f, r, err := ledongthucpdf.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	total := r.NumPage()
	from := opts.From
	if from <= 0 {
		from = 1
	}
	to := opts.To
	if to <= 0 || to > total {
		to = total
	}
	if from > total {
		return nil, fmt.Errorf("page %d out of range (1..%d)", from, total)
	}

	pages := make([]Page, 0, to-from+1)
	for i := from; i <= to; i++ {
		p := r.Page(i)
		if p.V.IsNull() {
			continue
		}
		text, err := p.GetPlainText(nil)
		if err != nil {
			text = ""
		}
		pages = append(pages, Page{
			Page:      i,
			Text:      text,
			WordCount: wordCount(text),
		})
		if opts.Limit > 0 && len(pages) >= opts.Limit {
			break
		}
	}
	return pages, nil
}

// AllText returns the concatenated plain text of every page, separated by form feeds.
// Used for the simple "give me everything" case to keep token count low.
func AllText(path string) (string, error) {
	f, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer func() { _ = f.Close() }()

	stat, err := f.Stat()
	if err != nil {
		return "", err
	}
	r, err := ledongthucpdf.NewReader(f, stat.Size())
	if err != nil {
		return "", err
	}

	b, err := r.GetPlainText()
	if err != nil {
		return "", err
	}
	buf := new(bytes.Buffer)
	if _, err := io.Copy(buf, b); err != nil {
		return "", err
	}
	return buf.String(), nil
}

// SearchOptions configure pdf search behavior.
type SearchOptions struct {
	Page          int  // 0 = all pages
	CaseSensitive bool // default: case-insensitive
	Context       int  // words of surrounding context (default 30)
	Limit         int  // max results (default 100)
}

// SearchResult is one keyword hit inside a PDF.
type SearchResult struct {
	Page    int    `json:"page"`
	Line    int    `json:"line"`    // 1-based line number within page text
	Snippet string `json:"snippet"` // surrounding context
	Offset  int    `json:"offset"`  // character offset within page text
}

// Search finds all occurrences of keyword in the PDF text content.
// Returns one SearchResult per hit with surrounding context.
func Search(path, keyword string, opts SearchOptions) ([]SearchResult, error) {
	if keyword == "" {
		return nil, errors.New("keyword cannot be empty")
	}

	limit := opts.Limit
	if limit <= 0 {
		limit = 100
	}
	ctxWords := opts.Context
	if ctxWords <= 0 {
		ctxWords = 30
	}

	ro := ReadOptions{}
	if opts.Page > 0 {
		ro.From = opts.Page
		ro.To = opts.Page
	}
	pages, err := Read(path, ro)
	if err != nil {
		return nil, err
	}

	needle := keyword
	if !opts.CaseSensitive {
		needle = strings.ToLower(keyword)
	}

	var results []SearchResult
	for _, pg := range pages {
		text := pg.Text
		searchText := text
		if !opts.CaseSensitive {
			searchText = strings.ToLower(text)
		}

		lines := strings.Split(text, "\n")
		offset := 0
		for lineNum, line := range lines {
			searchLine := line
			if !opts.CaseSensitive {
				searchLine = strings.ToLower(line)
			}
			if strings.Contains(searchLine, needle) {
				snippet := extractSnippet(line, keyword, ctxWords, opts.CaseSensitive)
				results = append(results, SearchResult{
					Page:    pg.Page,
					Line:    lineNum + 1,
					Snippet: snippet,
					Offset:  offset,
				})
				if len(results) >= limit {
					return results, nil
				}
			}
			offset += len(line) + 1 // +1 for the newline
		}

		// If no newlines in text, also try as a single block
		if len(lines) <= 1 && strings.Contains(searchText, needle) && (len(results) == 0 || results[len(results)-1].Page != pg.Page) {
			snippet := extractSnippet(text, keyword, ctxWords, opts.CaseSensitive)
			results = append(results, SearchResult{
				Page:    pg.Page,
				Line:    1,
				Snippet: snippet,
				Offset:  0,
			})
			if len(results) >= limit {
				return results, nil
			}
		}
	}
	return results, nil
}

// extractSnippet returns a window of words around the first match of keyword in text.
func extractSnippet(text, keyword string, ctxWords int, caseSensitive bool) string {
	words := strings.Fields(text)
	if len(words) == 0 {
		return ""
	}

	needle := keyword
	haystack := words
	if !caseSensitive {
		needle = strings.ToLower(keyword)
		lower := make([]string, len(words))
		for i, w := range words {
			lower[i] = strings.ToLower(w)
		}
		haystack = lower
	}

	// Find the first word that contains the keyword
	matchIdx := -1
	for i, w := range haystack {
		if strings.Contains(w, needle) {
			matchIdx = i
			break
		}
	}
	if matchIdx < 0 {
		// Keyword spans multiple words or not found in individual words
		// Return first N words as fallback
		end := ctxWords
		if end > len(words) {
			end = len(words)
		}
		return strings.Join(words[:end], " ")
	}

	start := matchIdx - ctxWords/2
	if start < 0 {
		start = 0
	}
	end := matchIdx + ctxWords/2 + 1
	if end > len(words) {
		end = len(words)
	}

	snippet := strings.Join(words[start:end], " ")
	if start > 0 {
		snippet = "..." + snippet
	}
	if end < len(words) {
		snippet = snippet + "..."
	}
	return snippet
}

// Info exposes the lightweight metadata pdfcpu reports for a file.
type Info struct {
	Path       string `json:"path"`
	Pages      int    `json:"pages"`
	Title      string `json:"title,omitempty"`
	Author     string `json:"author,omitempty"`
	Subject    string `json:"subject,omitempty"`
	Keywords   string `json:"keywords,omitempty"`
	Creator    string `json:"creator,omitempty"`
	Producer   string `json:"producer,omitempty"`
	Created    string `json:"created,omitempty"`
	Modified   string `json:"modified,omitempty"`
	Encrypted  bool   `json:"encrypted,omitempty"`
	Version    string `json:"version,omitempty"`
	SizeBytes  int64  `json:"sizeBytes"`
	PageSize   string `json:"pageSize,omitempty"`
	Tagged     bool   `json:"tagged,omitempty"`
	Linearized bool   `json:"linearized,omitempty"`
}

// PageCount returns the number of pages in the PDF.
func PageCount(path string) (int, error) {
	return pdfapi.PageCountFile(path)
}

// Merge concatenates the inputs into outPath, in order.
func Merge(inputs []string, outPath string) error {
	if len(inputs) < 2 {
		return errors.New("merge requires at least 2 input files")
	}
	for _, in := range inputs {
		if _, err := os.Stat(in); err != nil {
			return fmt.Errorf("input not accessible: %s: %w", in, err)
		}
	}
	return pdfapi.MergeCreateFile(inputs, outPath, false, nil)
}

// SplitEvery splits inPath into files of `span` pages each, written to outDir.
func SplitEvery(inPath, outDir string, span int) error {
	if span <= 0 {
		return errors.New("span must be > 0")
	}
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return err
	}
	return pdfapi.SplitFile(inPath, outDir, span, nil)
}

// Trim writes a new PDF containing only the listed pages (or page ranges).
//
// pages accepts comma-separated page numbers and ranges: "1,3,5-7,10".
// Pass "" to keep all pages (effectively a copy operation).
//
// Internally, each comma-separated token is passed as a separate element to
// pdfcpu, which expects "1", "3", "5-7", "10" rather than "1,3,5-7,10".
func Trim(inPath, outPath, pages string) error {
	sel := splitPageSpec(pages)
	return pdfapi.TrimFile(inPath, outPath, sel, nil)
}

// WatermarkText adds a text watermark to inPath, writing the result to outPath.
// description follows pdfcpu's WM syntax, e.g.:
//
//	"font:Helvetica, points:24, opacity:0.3, rotation:45"
func WatermarkText(inPath, outPath, text, description string) error {
	if text == "" {
		return errors.New("watermark text cannot be empty")
	}
	wm, err := pdfapi.TextWatermark(text, description, true, false, 1)
	if err != nil {
		return err
	}
	return pdfapi.AddWatermarksFile(inPath, outPath, nil, wm, nil)
}

// PageDimensions returns the width/height of every page in points.
func PageDimensions(path string) ([]PageSize, error) {
	dims, err := pdfapi.PageDimsFile(path)
	if err != nil {
		return nil, err
	}
	out := make([]PageSize, 0, len(dims))
	for i, d := range dims {
		out = append(out, PageSize{
			Page:   i + 1,
			Width:  d.Width,
			Height: d.Height,
		})
	}
	return out, nil
}

// PageSize is one page's dimensions in points (1pt = 1/72 inch).
type PageSize struct {
	Page   int     `json:"page"`
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
}

// wordCount counts whitespace-separated tokens in s.
func wordCount(s string) int {
	return len(strings.Fields(s))
}

// splitPageSpec splits a comma-separated page spec into individual tokens
// suitable for pdfcpu's page selection slice. Each comma-separated piece
// becomes a separate element (e.g. "1,3,5-7" → ["1","3","5-7"]).
// An empty input returns nil (meaning "all pages" in pdfcpu).
func splitPageSpec(s string) []string {
	if s == "" {
		return nil
	}
	parts := strings.Split(s, ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			out = append(out, p)
		}
	}
	if len(out) == 0 {
		return nil
	}
	return out
}

// ParsePageList validates a "1,3,5-7" style list. Returns the list as-is on success.
func ParsePageList(s string) (string, error) {
	if s == "" {
		return "", nil
	}
	for _, part := range strings.Split(s, ",") {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if strings.Contains(part, "-") {
			bounds := strings.SplitN(part, "-", 2)
			startStr := strings.TrimSpace(bounds[0])
			endStr := strings.TrimSpace(bounds[1])
			if startStr == "" {
				return "", fmt.Errorf("invalid page range: missing start in %q", part)
			}
			if _, err := strconv.Atoi(startStr); err != nil {
				return "", fmt.Errorf("invalid page range start: %q", bounds[0])
			}
			// "2-" means "from page 2 to end" — valid open-ended range.
			if endStr == "" {
				continue
			}
			if _, err := strconv.Atoi(endStr); err != nil {
				return "", fmt.Errorf("invalid page range end: %q", bounds[1])
			}
			continue
		}
		if _, err := strconv.Atoi(part); err != nil {
			return "", fmt.Errorf("invalid page number: %q", part)
		}
	}
	return s, nil
}

// FullInfo returns the consolidated metadata pdfcpu reports for a file.
//
// The struct is rich (page boundaries, encryption flags, attachments) which is
// exactly what AI Agents need to plan further actions; we expose it directly.
func FullInfo(path string) (map[string]any, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	info, err := pdfapi.PDFInfo(f, path, nil, false, nil)
	if err != nil {
		return nil, err
	}

	stat, _ := os.Stat(path)
	out := map[string]any{
		"path":             path,
		"sizeBytes":        statSize(stat),
		"version":          info.Version,
		"pageCount":        info.PageCount,
		"title":            info.Title,
		"author":           info.Author,
		"subject":          info.Subject,
		"keywords":         info.Keywords,
		"creator":          info.Creator,
		"producer":         info.Producer,
		"creationDate":     info.CreationDate,
		"modificationDate": info.ModificationDate,
		"encrypted":        info.Encrypted,
		"watermarked":      info.Watermarked,
		"linearized":       info.Linearized,
		"tagged":           info.Tagged,
		"form":             info.Form,
		"signatures":       info.Signatures,
		"bookmarks":        info.Outlines,
		"attachments":      info.Attachments,
	}
	if len(info.Dimensions) > 0 {
		dims := make([]map[string]float64, 0, len(info.Dimensions))
		for _, d := range info.Dimensions {
			dims = append(dims, map[string]float64{"width": d.Width, "height": d.Height})
		}
		out["dimensions"] = dims
	}
	return out, nil
}

func statSize(s os.FileInfo) int64 {
	if s == nil {
		return 0
	}
	return s.Size()
}

// Rotate writes a copy of inPath into outPath where the listed pages are rotated
// by `rotation` degrees (must be a multiple of 90: 90 / 180 / 270).
// pages="" rotates every page.
func Rotate(inPath, outPath string, rotation int, pages string) error {
	if rotation%90 != 0 {
		return fmt.Errorf("rotation must be a multiple of 90, got %d", rotation)
	}
	var sel []string
	if pages != "" {
		sel = []string{pages}
	}
	return pdfapi.RotateFile(inPath, outPath, rotation, sel, nil)
}

// Optimize re-writes the PDF with object reuse and cross-reference compaction.
// Typically reduces size by 5–30% with no visible quality change.
func Optimize(inPath, outPath string) error {
	return pdfapi.OptimizeFile(inPath, outPath, nil)
}

// Encrypt protects the PDF with a user password. ownerPW (admin) defaults to
// userPW when empty. The PDF is encrypted with the strongest algorithm pdfcpu
// supports for the file's PDF version.
func Encrypt(inPath, outPath, userPW, ownerPW string) error {
	if userPW == "" {
		return errors.New("user password cannot be empty")
	}
	if ownerPW == "" {
		ownerPW = userPW
	}
	conf := pdfmodel.NewDefaultConfiguration()
	conf.UserPW = userPW
	conf.OwnerPW = ownerPW
	conf.Cmd = pdfmodel.ENCRYPT
	return pdfapi.EncryptFile(inPath, outPath, conf)
}

// Decrypt strips encryption when the right password is provided. Either userPW
// or ownerPW is sufficient.
func Decrypt(inPath, outPath, userPW, ownerPW string) error {
	conf := pdfmodel.NewDefaultConfiguration()
	conf.UserPW = userPW
	conf.OwnerPW = ownerPW
	conf.Cmd = pdfmodel.DECRYPT
	return pdfapi.DecryptFile(inPath, outPath, conf)
}

// ExtractImages dumps every image embedded in the PDF into outDir.
// pages="" extracts from every page.
func ExtractImages(inPath, outDir, pages string) error {
	if err := os.MkdirAll(outDir, 0755); err != nil {
		return err
	}
	var sel []string
	if pages != "" {
		sel = []string{pages}
	}
	return pdfapi.ExtractImagesFile(inPath, outDir, sel, nil)
}

// ListExtractedImages walks outDir after ExtractImages and returns one entry per
// image written. Useful for AI Agents that want to immediately enumerate the
// produced files without reaching into the filesystem themselves.
func ListExtractedImages(outDir string) ([]ExtractedImage, error) {
	entries, err := os.ReadDir(outDir)
	if err != nil {
		return nil, err
	}
	var out []ExtractedImage
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		info, err := e.Info()
		if err != nil {
			continue
		}
		out = append(out, ExtractedImage{
			Name:  e.Name(),
			Path:  outDir + string(os.PathSeparator) + e.Name(),
			Bytes: info.Size(),
		})
	}
	return out, nil
}

// ExtractedImage mirrors common.ExtractedImage; defined here too so the pdf
// engine has a self-contained public API.
type ExtractedImage struct {
	Name  string `json:"name"`
	Path  string `json:"path"`
	Bytes int64  `json:"bytes"`
}

// ---------------------------------------------------------------------------
// New PDF engine functions: bookmarks, reorder, insert blank, stamp, metadata
// ---------------------------------------------------------------------------

// Bookmark is one entry in the PDF outline tree.
type Bookmark struct {
	Title    string     `json:"title"`
	Page     int        `json:"page"`
	Bold     bool       `json:"bold,omitempty"`
	Italic   bool       `json:"italic,omitempty"`
	Children []Bookmark `json:"kids,omitempty"`
}

// Bookmarks reads the outline/bookmarks tree from a PDF.
func Bookmarks(path string) ([]Bookmark, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	bms, err := pdfapi.Bookmarks(f, nil)
	if err != nil {
		return nil, err
	}
	return convertBookmarks(bms), nil
}

func convertBookmarks(bms []pdfcpu.Bookmark) []Bookmark {
	out := make([]Bookmark, 0, len(bms))
	for _, b := range bms {
		bm := Bookmark{
			Title:  b.Title,
			Page:   b.PageFrom,
			Bold:   b.Bold,
			Italic: b.Italic,
		}
		if len(b.Kids) > 0 {
			bm.Children = convertBookmarks(b.Kids)
		}
		out = append(out, bm)
	}
	return out
}

// BookmarkSpec is used to add bookmarks from JSON input.
type BookmarkSpec struct {
	Title    string         `json:"title"`
	Page     int            `json:"page"`
	Children []BookmarkSpec `json:"kids,omitempty"`
}

// AddBookmarks adds bookmarks to a PDF.
func AddBookmarks(inPath, outPath string, specs []BookmarkSpec, replace bool) error {
	bms := toPDFCPUBookmarks(specs)
	return pdfapi.AddBookmarksFile(inPath, outPath, bms, replace, nil)
}

func toPDFCPUBookmarks(specs []BookmarkSpec) []pdfcpu.Bookmark {
	out := make([]pdfcpu.Bookmark, 0, len(specs))
	for _, s := range specs {
		bm := pdfcpu.Bookmark{
			Title:    s.Title,
			PageFrom: s.Page,
		}
		if len(s.Children) > 0 {
			bm.Kids = toPDFCPUBookmarks(s.Children)
		}
		out = append(out, bm)
	}
	return out
}

// Reorder reorders pages of a PDF. order is a comma-separated page list like "3,1,2,4".
func Reorder(inPath, outPath, order string) error {
	sel := splitPageSpec(order)
	return pdfapi.CollectFile(inPath, outPath, sel, nil)
}

// InsertBlank inserts `count` blank pages after `afterPage` (1-based).
func InsertBlank(inPath, outPath string, afterPage, count int) error {
	if count <= 0 {
		return errors.New("count must be > 0")
	}
	// InsertPagesFile inserts one blank page per selected page.
	// To insert `count` blanks after `afterPage`, we insert one at a time.
	currentPage := afterPage
	tmpIn := inPath
	for i := 0; i < count; i++ {
		isLast := i == count-1
		out := outPath
		if !isLast {
			out = inPath + fmt.Sprintf(".tmp%d", i)
		}
		sel := []string{strconv.Itoa(currentPage)}
		if err := pdfapi.InsertPagesFile(tmpIn, out, sel, false, nil, nil); err != nil {
			return err
		}
		if !isLast {
			if tmpIn != inPath {
				_ = os.Remove(tmpIn)
			}
			tmpIn = out
		}
		currentPage++
	}
	// Clean up any remaining temp files
	for i := 0; i < count-1; i++ {
		_ = os.Remove(inPath + fmt.Sprintf(".tmp%d", i))
	}
	return nil
}

// StampImage stamps an image on selected pages. pages="" means all pages.
func StampImage(inPath, outPath, imagePath, pages string) error {
	var sel []string
	if pages != "" {
		sel = []string{pages}
	}
	return pdfapi.AddImageWatermarksFile(inPath, outPath, sel, false, imagePath, "", nil)
}

// MetaUpdate holds metadata fields to set on a PDF.
type MetaUpdate struct {
	Title    string `json:"title,omitempty"`
	Author   string `json:"author,omitempty"`
	Subject  string `json:"subject,omitempty"`
	Keywords string `json:"keywords,omitempty"`
}

// ---------------------------------------------------------------------------
// CreatePDF — hand-rolled PDF generation
// ---------------------------------------------------------------------------

// CreateSpec describes a PDF to create from scratch.
type CreateSpec struct {
	Pages      []PageSpec `json:"pages"`
	Title      string     `json:"title,omitempty"`
	Author     string     `json:"author,omitempty"`
	Paper      string     `json:"paper,omitempty"`      // "A4" (default), "Letter", or "WxH"
	Font       string     `json:"font,omitempty"`       // global font family: "Helvetica"(default), "Times", "Courier"
	FontSize   float64    `json:"fontSize,omitempty"`   // global font size (default 12)
	Align      string     `json:"align,omitempty"`      // global align: "left"(default), "center", "right"
	LineHeight float64    `json:"lineHeight,omitempty"` // global line-height multiplier (default 1.4)
	Margin     float64    `json:"margin,omitempty"`     // global margin in points (default 72 = 1 inch)
}

// PageSpec is one page of content to create.
type PageSpec struct {
	Text       string  `json:"text"`                 // plain text, newlines respected
	FontSize   float64 `json:"fontSize,omitempty"`   // page-level font size (overrides global)
	Bold       bool    `json:"bold,omitempty"`       // bold variant
	Italic     bool    `json:"italic,omitempty"`     // italic variant
	Color      string  `json:"color,omitempty"`      // "R G B" in 0-1 range, e.g. "1 0 0" for red
	Font       string  `json:"font,omitempty"`       // font family: "Helvetica"(default), "Times", "Courier"
	Align      string  `json:"align,omitempty"`      // "left"(default), "center", "right"
	LineHeight float64 `json:"lineHeight,omitempty"` // line-height multiplier (default 1.4)
	Margin     float64 `json:"margin,omitempty"`     // margin in points (default 72)
	Underline  bool    `json:"underline,omitempty"`  // underline each line
}

// mergePageDefaults fills zero-value PageSpec fields from CreateSpec globals.
func mergePageDefaults(pg PageSpec, global CreateSpec) PageSpec {
	if pg.FontSize <= 0 && global.FontSize > 0 {
		pg.FontSize = global.FontSize
	}
	if pg.Font == "" && global.Font != "" {
		pg.Font = global.Font
	}
	if pg.Align == "" && global.Align != "" {
		pg.Align = global.Align
	}
	if pg.LineHeight <= 0 && global.LineHeight > 0 {
		pg.LineHeight = global.LineHeight
	}
	if pg.Margin <= 0 && global.Margin > 0 {
		pg.Margin = global.Margin
	}
	return pg
}

// CreatePDF writes a new PDF to outPath based on the given spec.
// Supports multiple font families (Helvetica/Times/Courier), text alignment,
// configurable line height, margins, and per-page styling.
// This implementation hand-rolls PDF bytes to avoid external dependencies
// for basic PDF creation.
func CreatePDF(outPath string, spec CreateSpec) error {
	if len(spec.Pages) == 0 {
		return errors.New("at least one page is required")
	}

	width, height := resolvePaperSize(spec.Paper)

	// Merge global defaults into each page
	pages := make([]PageSpec, len(spec.Pages))
	for i, pg := range spec.Pages {
		pages[i] = mergePageDefaults(pg, spec)
	}

	// Collect all unique font names needed across all pages
	uniqueFonts := orderedStringSet{}
	fontPerLine := make([][]string, len(pages)) // fonts used per page
	for i, pg := range pages {
		mainFont := resolveFontName(pg.Font, pg.Bold, pg.Italic)
		uniqueFonts.Add(mainFont)
		fontPerLine[i] = []string{mainFont}
	}

	var body bytes.Buffer
	body.WriteString("%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")

	type objEntry struct {
		num    int
		offset int
	}
	var objects []objEntry

	addObj := func(num int, content string) {
		objects = append(objects, objEntry{num: num, offset: body.Len()})
		fmt.Fprintf(&body, "%d 0 obj\n%s\nendobj\n", num, content)
	}

	// Object numbering:
	//   1 = Catalog
	//   2 = Pages
	//   3..(2+N) = Page objects
	//   (3+N)..(2+2N) = Content stream objects
	//   (3+2N)..(2+2N+F) = Font objects (F = len(uniqueFonts))
	//   next = Info dict
	n := len(pages)
	pageObjStart := 3
	contentObjStart := pageObjStart + n
	fontObjStart := contentObjStart + n
	infoObjNum := fontObjStart + uniqueFonts.Len()

	// Build font name → object number + reference map
	fontObjMap := map[string]int{}    // fontName → objNum
	fontRefMap := map[string]string{} // fontName → "/F1", "/F2", ...
	for i, fname := range uniqueFonts.Items() {
		objNum := fontObjStart + i
		fontObjMap[fname] = objNum
		fontRefMap[fname] = fmt.Sprintf("/F%d", i+1)
	}

	// Build Kids array
	var kids string
	for i := range pages {
		if i > 0 {
			kids += " "
		}
		kids += fmt.Sprintf("%d 0 R", pageObjStart+i)
	}

	addObj(1, "<< /Type /Catalog /Pages 2 0 R >>")
	addObj(2, fmt.Sprintf("<< /Type /Pages /Kids [%s] /Count %d >>", kids, n))

	// Page objects and content streams
	for i, pg := range pages {
		pn := pageObjStart + i
		cn := contentObjStart + i

		// Build font resource map for this page
		usedFonts := map[string]bool{}
		for _, fname := range fontPerLine[i] {
			usedFonts[fname] = true
		}
		var fontMapParts []string
		for fname := range usedFonts {
			fontMapParts = append(fontMapParts, fmt.Sprintf("%s %d 0 R", fontRefMap[fname], fontObjMap[fname]))
		}
		fontMap := strings.Join(fontMapParts, " ")

		addObj(pn, fmt.Sprintf("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R /Resources << /Font << %s >> >> >>", int(width), int(height), cn, fontMap))

		mainFont := resolveFontName(pg.Font, pg.Bold, pg.Italic)
		fontRef := fontRefMap[mainFont]
		stream := buildTextStream(pg.Text, width, height, pg.FontSize, pg.Color, fontRef, mainFont, pg.Align, pg.LineHeight, pg.Margin, pg.Underline)
		addObj(cn, fmt.Sprintf("<< /Length %d >>\nstream\n%sendstream", len(stream), stream))
	}

	// Font objects
	for _, fname := range uniqueFonts.Items() {
		objNum := fontObjMap[fname]
		addObj(objNum, fmt.Sprintf("<< /Type /Font /Subtype /Type1 /BaseFont /%s >>", fname))
	}

	// Info dictionary (optional metadata)
	if spec.Title != "" || spec.Author != "" {
		var infoEntries string
		if spec.Title != "" {
			infoEntries += fmt.Sprintf(" /Title (%s)", escapePDFStr(spec.Title))
		}
		if spec.Author != "" {
			infoEntries += fmt.Sprintf(" /Author (%s)", escapePDFStr(spec.Author))
		}
		addObj(infoObjNum, "<<"+infoEntries+" >>")
	} else {
		addObj(infoObjNum, "<< >>")
	}

	// Cross-reference table
	xrefOffset := body.Len()
	total := infoObjNum + 1
	body.WriteString("xref\n")
	fmt.Fprintf(&body, "0 %d\n", total)
	body.WriteString("0000000000 65535 f \n")
	for i := 1; i < total; i++ {
		var found *objEntry
		for j := range objects {
			if objects[j].num == i {
				found = &objects[j]
				break
			}
		}
		if found == nil {
			body.WriteString("0000000000 00000 f \n")
			continue
		}
		fmt.Fprintf(&body, "%010d 00000 n \n", found.offset)
	}

	// Trailer
	trailer := fmt.Sprintf("trailer\n<< /Size %d /Root 1 0 R /Info %d 0 R >>\nstartxref\n%d\n%%%%EOF\n", total, infoObjNum, xrefOffset)
	body.WriteString(trailer)

	return os.WriteFile(outPath, body.Bytes(), 0644)
}

// orderedStringSet is a set of strings that preserves insertion order.
type orderedStringSet struct {
	items []string
	seen  map[string]bool
}

func (s *orderedStringSet) Add(v string) {
	if s.seen == nil {
		s.seen = map[string]bool{}
	}
	if !s.seen[v] {
		s.seen[v] = true
		s.items = append(s.items, v)
	}
}

func (s *orderedStringSet) Items() []string { return s.items }
func (s *orderedStringSet) Len() int        { return len(s.items) }

// resolvePaperSize returns (width, height) in points for a paper name.
// Default is US Letter (612 x 792).
func resolvePaperSize(paper string) (float64, float64) {
	switch strings.ToUpper(strings.TrimSpace(paper)) {
	case "A4":
		return 595.28, 841.89
	case "LETTER", "":
		return 612, 792
	default:
		// Try "WxH" format
		if idx := strings.Index(strings.ToLower(paper), "x"); idx > 0 {
			w, err1 := strconv.ParseFloat(strings.TrimSpace(paper[:idx]), 64)
			h, err2 := strconv.ParseFloat(strings.TrimSpace(paper[idx+1:]), 64)
			if err1 == nil && err2 == nil && w > 0 && h > 0 {
				return w, h
			}
		}
		return 612, 792 // fallback to Letter
	}
}

// buildTextStream generates a PDF content stream that renders plain text.
// Supports font selection, text alignment, configurable line height, margins,
// and underline decoration.
func buildTextStream(text string, width, height, fontSize float64, color, fontRef, fontName, align string, lineHeightMul, marginPt float64, underline bool) string {
	if fontSize <= 0 {
		fontSize = 12.0
	}
	if lineHeightMul <= 0 {
		lineHeightMul = 1.4
	}
	if marginPt <= 0 {
		marginPt = 72.0
	}

	cwf := charWidthFactor(fontName)
	lineHeight := fontSize * lineHeightMul
	maxX := width - 2*marginPt
	charsPerLine := int(maxX / (fontSize * cwf))
	if charsPerLine < 1 {
		charsPerLine = 1
	}

	lines := wrapText(text, charsPerLine)

	var buf bytes.Buffer
	y := height - marginPt
	buf.WriteString("BT\n")

	if color != "" {
		fmt.Fprintf(&buf, "%s rg\n", color)
	}
	fmt.Fprintf(&buf, "%s %.1f Tf\n", fontRef, fontSize)

	// Track underline positions (absolute coords)
	type ulSeg struct{ x1, x2, y float64 }
	var ulSegs []ulSeg

	prevX := marginPt
	for i, line := range lines {
		escaped := escapePDFStr(line)
		lineWidth := float64(len(line)) * fontSize * cwf

		var x float64
		switch strings.ToLower(strings.TrimSpace(align)) {
		case "center":
			x = (width - lineWidth) / 2
			if x < marginPt {
				x = marginPt
			}
		case "right":
			x = width - marginPt - lineWidth
			if x < marginPt {
				x = marginPt
			}
		default:
			x = marginPt
		}

		if i == 0 {
			fmt.Fprintf(&buf, "%.1f %.1f Td (%s) Tj\n", x, y, escaped)
		} else {
			fmt.Fprintf(&buf, "%.1f -%.1f Td (%s) Tj\n", x-prevX, lineHeight, escaped)
		}
		prevX = x

		if underline {
			baseY := y - float64(i)*lineHeight
			ulSegs = append(ulSegs, ulSeg{x1: x, x2: x + lineWidth, y: baseY - fontSize*0.15})
		}
	}

	buf.WriteString("ET\n")

	// Draw underlines after text block (strokes are outside BT/ET)
	if len(ulSegs) > 0 {
		thickness := fontSize * 0.05
		if thickness < 0.5 {
			thickness = 0.5
		}
		fmt.Fprintf(&buf, "%.1f w\n", thickness)
		for _, ul := range ulSegs {
			fmt.Fprintf(&buf, "%.1f %.1f m %.1f %.1f l S\n", ul.x1, ul.y, ul.x2, ul.y)
		}
	}

	return buf.String()
}

// wrapText splits text into lines, wrapping at maxChars per line.
func wrapText(text string, maxChars int) []string {
	if maxChars <= 0 {
		maxChars = 80
	}
	var result []string
	for _, paragraph := range strings.Split(text, "\n") {
		if len(paragraph) == 0 {
			result = append(result, "")
			continue
		}
		words := strings.Fields(paragraph)
		line := ""
		for _, w := range words {
			if line == "" {
				line = w
				continue
			}
			if len(line)+1+len(w) > maxChars {
				result = append(result, line)
				line = w
			} else {
				line += " " + w
			}
		}
		if line != "" {
			result = append(result, line)
		}
	}
	return result
}

// escapePDFStr escapes characters special in PDF string literals.
func escapePDFStr(s string) string {
	var b strings.Builder
	for _, r := range s {
		switch r {
		case '(', ')', '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		default:
			if r > 0x7f {
				b.WriteByte('?')
			} else {
				b.WriteRune(r)
			}
		}
	}
	return b.String()
}

// ---------------------------------------------------------------------------
// Font helpers — PDF 14 standard base fonts
// ---------------------------------------------------------------------------

// resolveFontName returns the standard PDF font name for a given family + style.
// Supported families: "Helvetica" (default/sans), "Times" (serif), "Courier" (mono).
// Aliases: "serif" → Times, "sans" → Helvetica, "mono"/"monospace" → Courier.
func resolveFontName(family string, bold, italic bool) string {
	switch strings.ToLower(strings.TrimSpace(family)) {
	case "times", "serif":
		switch {
		case bold && italic:
			return "Times-BoldItalic"
		case bold:
			return "Times-Bold"
		case italic:
			return "Times-Italic"
		default:
			return "Times-Roman"
		}
	case "courier", "mono", "monospace":
		switch {
		case bold && italic:
			return "Courier-BoldOblique"
		case bold:
			return "Courier-Bold"
		case italic:
			return "Courier-Oblique"
		default:
			return "Courier"
		}
	default: // helvetica, sans, sans-serif, ""
		switch {
		case bold && italic:
			return "Helvetica-BoldOblique"
		case bold:
			return "Helvetica-Bold"
		case italic:
			return "Helvetica-Oblique"
		default:
			return "Helvetica"
		}
	}
}

// charWidthFactor returns the average character width as a fraction of font size.
// Used for line wrapping and alignment calculations.
func charWidthFactor(fontName string) float64 {
	lower := strings.ToLower(fontName)
	if strings.Contains(lower, "courier") {
		return 0.60 // monospace — exact
	}
	if strings.Contains(lower, "times") {
		return 0.45 // serif — narrower
	}
	return 0.50 // helvetica (default sans)
}

// ---------------------------------------------------------------------------
// ReplaceText — content stream manipulation
// ---------------------------------------------------------------------------

// Replacement is a find/replace pair.
type Replacement struct {
	Find    string `json:"find"`
	Replace string `json:"replace"`
}

// ReplaceText performs find/replace on PDF text content stream literals.
//
// It accesses the content stream for each page via pdfcpu's XRefTable,
// finds all parenthesized string literals in text-showing operators (Tj, ', "),
// and replaces matching text. After modification the streams are re-encoded and
// the PDF is written back through pdfcpu's WriteContext.
//
// Limitations:
//   - Only replaces text stored as parenthesized string literals (not hex <...>)
//   - Text split across multiple Tj operators is not matched
//   - Font encoding may cause some characters to not match as expected
//
// Returns the total number of replacements made across all pages.
func ReplaceText(inPath, outPath string, pairs []Replacement) (int, error) {
	if len(pairs) == 0 {
		return 0, errors.New("at least one replacement pair is required")
	}

	src, err := os.Open(inPath)
	if err != nil {
		return 0, err
	}
	defer func() { _ = src.Close() }()

	conf := pdfmodel.NewDefaultConfiguration()
	ctx, err := pdfapi.ReadAndValidate(src, conf)
	if err != nil {
		return 0, fmt.Errorf("reading PDF: %w", err)
	}

	totalHits := 0

	// Walk the XRefTable looking for content streams.
	// We identify content streams by checking for text-showing operators (Tj)
	// in the stream data, regardless of the IsPageContent flag (which may not
	// be set for all PDFs, especially hand-crafted ones).
	for _, entry := range ctx.XRefTable.Table {
		if entry == nil || entry.Object == nil {
			continue
		}
		sd, ok := entry.Object.(pdftypes.StreamDict)
		if !ok {
			continue
		}

		// Get the stream text — prefer Content (decoded), fall back to Raw.
		streamBytes := sd.Content
		if len(streamBytes) == 0 {
			streamBytes = sd.Raw
		}
		if len(streamBytes) == 0 {
			continue
		}

		// Quick check: is this a content stream? Look for Tj operator.
		if !isLikelyContentStream(streamBytes) {
			continue
		}

		stream := string(streamBytes)
		modified := stream
		for _, p := range pairs {
			modified = replaceInContentStream(modified, p.Find, p.Replace, &totalHits)
		}

		if modified != stream {
			sd.Content = []byte(modified)
			if err := sd.Encode(); err != nil {
				// Non-fatal: skip this stream
				continue
			}
			entry.Object = sd
		}
	}

	out, err := os.Create(outPath)
	if err != nil {
		return 0, err
	}
	defer func() { _ = out.Close() }()

	if err := pdfapi.WriteContext(ctx, out); err != nil {
		return 0, err
	}
	return totalHits, nil
}

// isLikelyContentStream checks if bytes look like a page content stream
// by looking for common PDF text operators.
func isLikelyContentStream(b []byte) bool {
	s := string(b)
	return strings.Contains(s, " Tj") || strings.Contains(s, "Tj\n") ||
		strings.Contains(s, "BT ") || strings.Contains(s, "BT\n")
}

// replaceInContentStream finds and replaces text inside PDF text-showing operators.
// It handles parenthesized strings: (text) Tj and hex strings: <hex> Tj.
// Returns the modified stream and increments *hits for each replacement.
func replaceInContentStream(stream, find, replace string, hits *int) string {
	if find == "" {
		return stream
	}

	var result strings.Builder
	i := 0
	for i < len(stream) {
		// Parenthesized string literal: (...)
		if stream[i] == '(' {
			end := findClosingParen(stream, i)
			if end > i {
				inner := stream[i+1 : end]
				unescaped := unescapePDFString(inner)
				if strings.Contains(unescaped, find) {
					newText := strings.ReplaceAll(unescaped, find, replace)
					*hits += strings.Count(unescaped, find)
					result.WriteByte('(')
					result.WriteString(escapePDFLiteralChars(newText))
					result.WriteByte(')')
					i = end + 1
					continue
				}
			}
		}
		// Hex string literal: <hex>
		if stream[i] == '<' && i+1 < len(stream) && stream[i+1] != '<' {
			end := strings.IndexByte(stream[i:], '>')
			if end > 0 {
				end += i
				hexStr := stream[i+1 : end]
				decoded := hexDecode(hexStr)
				if decoded != "" && strings.Contains(decoded, find) {
					newText := strings.ReplaceAll(decoded, find, replace)
					*hits += strings.Count(decoded, find)
					result.WriteByte('<')
					result.WriteString(hexEncode(newText))
					result.WriteByte('>')
					i = end + 1
					continue
				}
			}
		}
		result.WriteByte(stream[i])
		i++
	}
	return result.String()
}

// hexDecode converts a PDF hex string (e.g. "48656C6C6F") to text (PDFDocEncoding).
func hexDecode(hex string) string {
	hex = strings.Join(strings.Fields(hex), "") // remove whitespace
	if len(hex)%2 != 0 {
		hex += "0" // pad odd-length
	}
	out := make([]byte, 0, len(hex)/2)
	for i := 0; i < len(hex); i += 2 {
		b, err := strconv.ParseUint(hex[i:i+2], 16, 8)
		if err != nil {
			return ""
		}
		out = append(out, byte(b))
	}
	return string(out)
}

// hexEncode converts text to a PDF hex string.
func hexEncode(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		fmt.Fprintf(&b, "%02X", s[i])
	}
	return b.String()
}

// findClosingParen finds the matching ')' for a '(' at position start in s.
// Handles escaped characters (\(, \), \\).
func findClosingParen(s string, start int) int {
	depth := 0
	for i := start; i < len(s); i++ {
		if s[i] == '\\' && i+1 < len(s) {
			i++ // skip escaped char
			continue
		}
		if s[i] == '(' {
			depth++
		} else if s[i] == ')' {
			depth--
			if depth == 0 {
				return i
			}
		}
	}
	return -1
}

// unescapePDFString removes PDF string escape sequences.
func unescapePDFString(s string) string {
	var b strings.Builder
	for i := 0; i < len(s); i++ {
		if s[i] == '\\' && i+1 < len(s) {
			i++
			switch s[i] {
			case 'n':
				b.WriteByte('\n')
			case 'r':
				b.WriteByte('\r')
			case 't':
				b.WriteByte('\t')
			case 'b':
				b.WriteByte('\b')
			case 'f':
				b.WriteByte('\f')
			default:
				b.WriteByte(s[i])
			}
		} else {
			b.WriteByte(s[i])
		}
	}
	return b.String()
}

// escapePDFLiteralChars escapes characters that are special inside PDF string literals.
func escapePDFLiteralChars(s string) string {
	var b strings.Builder
	for _, r := range s {
		switch r {
		case '(', ')', '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		case '\n':
			b.WriteString("\\n")
		case '\r':
			b.WriteString("\\r")
		case '\t':
			b.WriteString("\\t")
		default:
			b.WriteRune(r)
		}
	}
	return b.String()
}

// SetMeta updates PDF metadata. Empty fields are left unchanged.
//
// Implementation: reads the full PDF context via ReadAndValidate, modifies the
// Info dictionary object directly (Title, Author, Subject, Keywords), then
// writes it back via WriteContext. Setting ctx.XRefTable.Title/Author/etc. does
// NOT work because pdfcpu's ensureInfoDict only writes Producer and dates; the
// Info dict entries must be patched in place.
func SetMeta(inPath, outPath string, meta MetaUpdate) error {
	src, err := os.Open(inPath)
	if err != nil {
		return err
	}
	defer func() { _ = src.Close() }()

	conf := pdfmodel.NewDefaultConfiguration()
	ctx, err := pdfapi.ReadAndValidate(src, conf)
	if err != nil {
		// Fall back to simple copy if read-and-validate is not available
		return pdfapi.CollectFile(inPath, outPath, nil, nil)
	}

	// Patch the Info dictionary directly. This is the only reliable way to
	// persist metadata through pdfcpu's write cycle (ensureInfoDict only
	// touches Producer / CreationDate / ModDate).
	//
	// If the PDF has no Info dict yet (ctx.Info == nil), create one so our
	// metadata survives the write cycle. ensureInfoDict will see it exists
	// and merge Producer/dates without destroying our entries.
	if ctx.Info == nil {
		d := pdftypes.NewDict()
		ir, err := ctx.IndRefForNewObject(d)
		if err == nil {
			ctx.Info = ir
		}
	}
	if ctx.Info != nil {
		d, err := ctx.DereferenceDict(*ctx.Info)
		if err == nil && d != nil {
			if meta.Title != "" {
				d.Update("Title", pdftypes.StringLiteral(meta.Title))
			}
			if meta.Author != "" {
				d.Update("Author", pdftypes.StringLiteral(meta.Author))
			}
			if meta.Subject != "" {
				d.Update("Subject", pdftypes.StringLiteral(meta.Subject))
			}
			if meta.Keywords != "" {
				d.Update("Keywords", pdftypes.StringLiteral(meta.Keywords))
			}
		}
	}

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}
	defer func() { _ = out.Close() }()

	return pdfapi.WriteContext(ctx, out)
}

// ---------------------------------------------------------------------------
// AddText — overlay text at specific coordinates on existing PDF pages
// ---------------------------------------------------------------------------

// TextOverlay describes a text annotation to add to PDF pages.
type TextOverlay struct {
	Text      string  `json:"text"`
	X         float64 `json:"x"`                  // x coordinate in points (from left)
	Y         float64 `json:"y"`                  // y coordinate in points (from bottom)
	FontSize  float64 `json:"fontSize,omitempty"` // default 12
	Bold      bool    `json:"bold,omitempty"`
	Italic    bool    `json:"italic,omitempty"`
	Underline bool    `json:"underline,omitempty"`
	Color     string  `json:"color,omitempty"` // "R G B" 0-1, e.g. "1 0 0" = red
	Font      string  `json:"font,omitempty"`  // font family: "Helvetica"(default), "Times", "Courier"
	Pages     string  `json:"pages,omitempty"` // "" = all pages; "1,3,5-7" etc.
}

// AddText overlays text at specific coordinates on existing PDF pages.
// It walks the XRefTable to find content streams, prepends text overlay
// operators at the beginning of each matching page's content stream, then
// writes the result via pdfcpu's WriteContext.
// Supports font family, bold, italic, underline, and color per overlay.
func AddText(inPath, outPath string, overlays []TextOverlay) error {
	if len(overlays) == 0 {
		return errors.New("at least one text overlay is required")
	}

	src, err := os.Open(inPath)
	if err != nil {
		return err
	}
	defer func() { _ = src.Close() }()

	conf := pdfmodel.NewDefaultConfiguration()
	ctx, err := pdfapi.ReadAndValidate(src, conf)
	if err != nil {
		return fmt.Errorf("reading PDF: %w", err)
	}

	// Collect unique font names needed by overlays
	uniqueFonts := orderedStringSet{}
	for _, ov := range overlays {
		uniqueFonts.Add(resolveFontName(ov.Font, ov.Bold, ov.Italic))
	}

	// Find existing font objects in the PDF and build a fontRefMap.
	// Also track the highest object number so we can add new ones.
	existingFontRefs := findExistingFonts(ctx)
	fontRefMap := map[string]string{}
	nextObjNum := 0
	for _, entry := range ctx.XRefTable.Table {
		if entry != nil {
			nextObjNum++ // crude: just count
		}
	}
	_ = nextObjNum

	for _, fname := range uniqueFonts.Items() {
		if ref, ok := existingFontRefs[fname]; ok {
			fontRefMap[fname] = ref
		} else {
			// Use /F9 as the overlay font reference to avoid clashing with existing fonts
			fontRefMap[fname] = "/F9"
		}
	}

	// Build page content map
	pageContentMap := buildPageContentMap(ctx)

	// For each content stream that belongs to a page, prepend overlay text
	for objNum, entry := range ctx.XRefTable.Table {
		if entry == nil || entry.Object == nil {
			continue
		}
		sd, ok := entry.Object.(pdftypes.StreamDict)
		if !ok {
			continue
		}

		streamBytes := sd.Content
		if len(streamBytes) == 0 {
			streamBytes = sd.Raw
		}
		if len(streamBytes) == 0 {
			continue
		}

		if !isLikelyContentStream(streamBytes) {
			continue
		}

		pageNum := pageContentMap[objNum]

		var applicable []TextOverlay
		for _, ov := range overlays {
			if pageNum > 0 && ov.Pages != "" {
				if !pageMatches(pageNum, ov.Pages) {
					continue
				}
			}
			applicable = append(applicable, ov)
		}
		if len(applicable) == 0 {
			continue
		}

		overlayStream := buildOverlayStream(applicable, fontRefMap)
		if overlayStream == "" {
			continue
		}

		original := string(streamBytes)
		// Append overlay after original content to avoid corrupting existing operators
		modified := original + "\n" + overlayStream

		sd.Content = []byte(modified)
		if err := sd.Encode(); err != nil {
			continue
		}
		entry.Object = sd
	}

	// Inject font objects for any overlay fonts not already in the PDF.
	// We add them as new XRefTable entries. For Helvetica (the default), most
	// PDFs already have it; for Times/Courier, we add a new font dict.
	for _, fname := range uniqueFonts.Items() {
		if _, ok := existingFontRefs[fname]; ok {
			continue // already in PDF
		}
		// Create a new font dictionary
		fontDict := pdftypes.Dict(map[string]pdftypes.Object{
			"Type":     pdftypes.Name("Font"),
			"Subtype":  pdftypes.Name("Type1"),
			"BaseFont": pdftypes.Name(fname),
		})
		ir, err := ctx.IndRefForNewObject(fontDict)
		if err == nil {
			// Store the indirect ref so pages can reference it
			fontRefMap[fname] = fmt.Sprintf("%d 0 R", ir.ObjectNumber)
			// Also update page Resources for pages that need this font
			addFontToPageResources(ctx, pageContentMap, fontRefMap[fname], fname)
		}
	}

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}
	defer func() { _ = out.Close() }()

	return pdfapi.WriteContext(ctx, out)
}

// findExistingFonts scans the XRefTable for existing Font dictionaries and
// returns a map from font name (e.g. "Helvetica") to its reference string.
func findExistingFonts(ctx *pdfmodel.Context) map[string]string {
	m := map[string]string{}
	for i, entry := range ctx.XRefTable.Table {
		if entry == nil || entry.Object == nil {
			continue
		}
		d, ok := entry.Object.(pdftypes.Dict)
		if !ok {
			continue
		}
		typeObj := d["Type"]
		if typeObj == nil {
			continue
		}
		if name, ok := typeObj.(pdftypes.Name); !ok || string(name) != "Font" {
			continue
		}
		baseFont := d["BaseFont"]
		if baseFont == nil {
			continue
		}
		if fname, ok := baseFont.(pdftypes.Name); ok {
			m[string(fname)] = fmt.Sprintf("%d 0 R", i)
		}
	}
	return m
}

// addFontToPageResources adds a font reference to each page's Resources/Font dict.
func addFontToPageResources(ctx *pdfmodel.Context, pageContentMap map[int]int, fontRef, fontName string) {
	for pageObjNum, entry := range ctx.XRefTable.Table {
		if entry == nil || entry.Object == nil {
			continue
		}
		d, ok := entry.Object.(pdftypes.Dict)
		if !ok {
			continue
		}
		typeObj := d["Type"]
		if typeObj == nil {
			continue
		}
		if name, ok := typeObj.(pdftypes.Name); !ok || string(name) != "Page" {
			continue
		}
		// Check if this page has content streams that we're modifying
		hasContent := false
		for contentObj := range pageContentMap {
			if pageContentMap[contentObj] == pageObjNum {
				// Check if this page's /Contents references this obj
				contents := d["Contents"]
				if ir, ok := contents.(pdftypes.IndirectRef); ok && int(ir.ObjectNumber) == contentObj {
					hasContent = true
					break
				}
			}
		}
		if !hasContent {
			continue
		}

		// Add font to Resources/Font
		resources := d["Resources"]
		if resources == nil {
			// Create Resources dict
			fontDict := pdftypes.Dict(map[string]pdftypes.Object{
				"F9": pdftypes.Name(fontRef), // simplified — will be resolved by pdfcpu
			})
			d["Resources"] = pdftypes.Dict(map[string]pdftypes.Object{
				"Font": fontDict,
			})
		} else if resDict, ok := resources.(pdftypes.Dict); ok {
			fonts := resDict["Font"]
			if fonts == nil {
				resDict["Font"] = pdftypes.Dict(map[string]pdftypes.Object{})
				fonts = resDict["Font"]
			}
			if fontDict, ok := fonts.(pdftypes.Dict); ok {
				fontDict["F9"] = pdftypes.Name(fontRef)
			}
		}
		entry.Object = d
	}
}

// buildPageContentMap builds a map from content-stream-object-number → page-number.
// It does this by finding Page dictionaries and extracting their /Contents reference.
func buildPageContentMap(ctx *pdfmodel.Context) map[int]int {
	m := make(map[int]int)
	pageNum := 0
	for _, entry := range ctx.XRefTable.Table {
		if entry == nil || entry.Object == nil {
			continue
		}
		d, ok := entry.Object.(pdftypes.Dict)
		if !ok {
			continue
		}
		// Check if this is a Page dictionary
		typeObj := d["Type"]
		if typeObj == nil {
			continue
		}
		name, ok := typeObj.(pdftypes.Name)
		if !ok || string(name) != "Page" {
			continue
		}
		pageNum++

		// Get /Contents — can be an indirect ref or an array
		contents := d["Contents"]
		if contents == nil {
			continue
		}
		if ir, ok := contents.(pdftypes.IndirectRef); ok {
			objNum := int(ir.ObjectNumber)
			m[objNum] = pageNum
		}
	}
	return m
}

// pageMatches checks if pageNum falls within a page spec like "1,3,5-7".
func pageMatches(pageNum int, spec string) bool {
	for _, part := range strings.Split(spec, ",") {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if strings.Contains(part, "-") {
			bounds := strings.SplitN(part, "-", 2)
			start, err1 := strconv.Atoi(strings.TrimSpace(bounds[0]))
			end, err2 := strconv.Atoi(strings.TrimSpace(bounds[1]))
			if err1 != nil || err2 != nil {
				continue
			}
			if pageNum >= start && pageNum <= end {
				return true
			}
		} else {
			n, err := strconv.Atoi(part)
			if err != nil {
				continue
			}
			if pageNum == n {
				return true
			}
		}
	}
	return false
}

// buildOverlayStream creates a PDF content stream snippet that renders text overlays.
// Supports font family, bold, italic, underline, and color per overlay.
// fontRefMap maps font names (e.g. "Helvetica-Bold") to PDF refs (e.g. "/F1").
// If nil, defaults to "/F1".
func buildOverlayStream(overlays []TextOverlay, fontRefMap map[string]string) string {
	var buf bytes.Buffer
	for _, ov := range overlays {
		fontSize := ov.FontSize
		if fontSize <= 0 {
			fontSize = 12.0
		}
		fontName := resolveFontName(ov.Font, ov.Bold, ov.Italic)
		fontRef := "/F1"
		if fontRefMap != nil {
			if ref, ok := fontRefMap[fontName]; ok {
				fontRef = ref
			}
		}
		escaped := escapePDFStr(ov.Text)
		cwf := charWidthFactor(fontName)
		lineWidth := float64(len(ov.Text)) * fontSize * cwf

		buf.WriteString("BT\n")
		if ov.Color != "" {
			fmt.Fprintf(&buf, "%s rg\n", ov.Color)
		}
		fmt.Fprintf(&buf, "%s %.1f Tf\n", fontRef, fontSize)
		fmt.Fprintf(&buf, "%.1f %.1f Td (%s) Tj\n", ov.X, ov.Y, escaped)
		buf.WriteString("ET\n")

		// Underline
		if ov.Underline {
			thickness := fontSize * 0.05
			if thickness < 0.5 {
				thickness = 0.5
			}
			ulY := ov.Y - fontSize*0.15
			fmt.Fprintf(&buf, "%.1f w %.1f %.1f m %.1f %.1f l S\n",
				thickness, ov.X, ulY, ov.X+lineWidth, ulY)
		}
	}
	return buf.String()
}
