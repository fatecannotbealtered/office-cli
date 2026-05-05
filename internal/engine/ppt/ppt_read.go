package ppt

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

// ReadSlides walks every ppt/slides/slide*.xml and returns one Slide per file,
// sorted by presentation order (sldIdLst in presentation.xml). Falls back to
// numeric file order if presentation.xml is missing or unparseable.
// Notes (if present) are read from notesSlideN.xml.
func ReadSlides(path string) ([]Slide, error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = zc.Close() }()

	byFileNum := map[int]slideEntry{}
	var fileNums []int
	notes := map[int]string{}

	for _, f := range zc.File {
		switch {
		case strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml"):
			n := slideNum(f.Name, "ppt/slides/slide")
			if n > 0 {
				se := slideEntry{fileNum: n, entryName: f.Name}
				byFileNum[n] = se
				fileNums = append(fileNums, n)
			}
		case strings.HasPrefix(f.Name, "ppt/notesSlides/notesSlide") && strings.HasSuffix(f.Name, ".xml"):
			n := slideNum(f.Name, "ppt/notesSlides/notesSlide")
			if n <= 0 {
				continue
			}
			data, err := common.ReadEntry(&zc.Reader, f.Name)
			if err != nil {
				continue
			}
			notes[n] = extractAllText(data)
		}
	}

	// Determine slide order: prefer sldIdLst from presentation.xml.
	// This respects ppt reorder operations.
	orderedFileNums := slideOrderFromPres(zc, byFileNum, fileNums)

	out := make([]Slide, 0, len(orderedFileNums))
	for i, fileNum := range orderedFileNums {
		se, ok := byFileNum[fileNum]
		if !ok {
			continue
		}
		data, err := common.ReadEntry(&zc.Reader, se.entryName)
		if err != nil {
			continue
		}
		title, bullets, allText := parseSlide(data)
		out = append(out, Slide{
			Index:   i + 1, // presentation position (1-based)
			File:    se.entryName,
			Title:   title,
			Bullets: bullets,
			Notes:   notes[fileNum],
			Text:    allText,
		})
	}
	return out, nil
}

// slideOrderFromPres reads presentation.xml's sldIdLst to determine the
// presentation-order list of slide file numbers. Falls back to sorted fileNums
// if the presentation.xml cannot be read or parsed.
func slideOrderFromPres(zc *zip.ReadCloser, byFileNum map[int]slideEntry, fileNums []int) []int {
	// Try to parse sldIdLst from presentation.xml
	presData, err := common.ReadEntry(&zc.Reader, "ppt/presentation.xml")
	if err != nil {
		sort.Ints(fileNums)
		return fileNums
	}
	sldEntries, _, err := parseSldIDLst(presData)
	if err != nil || len(sldEntries) == 0 {
		sort.Ints(fileNums)
		return fileNums
	}

	// Parse presentation.xml.rels to get rId → slide file number
	relsData, err := common.ReadEntry(&zc.Reader, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		sort.Ints(fileNums)
		return fileNums
	}
	var rels xmlRels
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		sort.Ints(fileNums)
		return fileNums
	}

	slideRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	rIDToFileNum := map[string]int{}
	for _, rel := range rels.Entries {
		if rel.Type != slideRelType {
			continue
		}
		num := slideNum("ppt/"+rel.Target, "ppt/slides/slide")
		if num > 0 {
			rIDToFileNum[rel.ID] = num
		}
	}

	// Build ordered list from sldIdLst entries
	var ordered []int
	seen := map[int]bool{}
	for _, e := range sldEntries {
		fileNum, ok := rIDToFileNum[e.RId]
		if !ok {
			continue
		}
		if _, exists := byFileNum[fileNum]; !exists {
			continue
		}
		if !seen[fileNum] {
			ordered = append(ordered, fileNum)
			seen[fileNum] = true
		}
	}
	// Append any slides not in sldIdLst (shouldn't happen, but be safe)
	sort.Ints(fileNums)
	for _, n := range fileNums {
		if !seen[n] {
			ordered = append(ordered, n)
		}
	}
	if len(ordered) == 0 {
		sort.Ints(fileNums)
		return fileNums
	}
	return ordered
}

// parseSlide returns (title, bullets, allText) extracted from a slide's XML bytes.
// Title detection: first shape whose placeholder type is "title" or "ctrTitle";
// fall back to the first non-empty paragraph in the slide.
func parseSlide(data []byte) (string, []string, string) {
	var s xmlSlide
	if err := xml.Unmarshal(data, &s); err != nil {
		return "", nil, ""
	}

	var title string
	var bullets []string
	var allParts []string

	for _, sh := range s.CSld.SpTree.Shapes {
		paragraphs := readShapeParagraphs(sh.TxBody)
		if len(paragraphs) == 0 {
			continue
		}
		isTitle := sh.NvSpPr.NvPr.Ph.Type == "title" || sh.NvSpPr.NvPr.Ph.Type == "ctrTitle"
		if isTitle && title == "" {
			title = strings.Join(paragraphs, " ")
			allParts = append(allParts, paragraphs...)
			continue
		}
		bullets = append(bullets, paragraphs...)
		allParts = append(allParts, paragraphs...)
	}

	for _, gf := range s.CSld.SpTree.GFrame {
		for _, tr := range gf.Graphic.GraphicData.Table.TR {
			for _, tc := range tr.TC {
				cellText := readShapeParagraphs(tc.TxBody)
				bullets = append(bullets, cellText...)
				allParts = append(allParts, cellText...)
			}
		}
	}

	if title == "" && len(allParts) > 0 {
		title = allParts[0]
	}
	return title, bullets, strings.Join(allParts, "\n")
}

// readShapeParagraphs flattens a txBody into a list of non-empty paragraph strings.
func readShapeParagraphs(tx xmlTxBody) []string {
	var out []string
	for _, p := range tx.Paragraphs {
		var sb strings.Builder
		for _, r := range p.Runs {
			sb.WriteString(r.Text)
		}
		if t := strings.TrimSpace(sb.String()); t != "" {
			out = append(out, t)
		}
	}
	return out
}

// extractAllText is a generic "give me every <a:t> value" helper used for notes.
func extractAllText(data []byte) string {
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	var sb strings.Builder
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "t" {
				var s string
				if err := dec.DecodeElement(&s, &t); err == nil {
					sb.WriteString(s)
					sb.WriteString(" ")
				}
			}
		}
	}
	return strings.TrimSpace(sb.String())
}

// slideNum extracts the trailing integer from "<prefix>N.xml". Returns 0 on failure.
func slideNum(name, prefix string) int {
	tail := strings.TrimPrefix(name, prefix)
	tail = strings.TrimSuffix(tail, ".xml")
	n, err := strconv.Atoi(tail)
	if err != nil {
		return 0
	}
	return n
}

// ReplaceText rewrites every slide XML by substituting each find with its replace.
// outPath is created (or overwritten); the source is never modified.
//
// Limitation: same as word.ReplaceText — matches must live within a single <a:t>
// element. Replacements spanning runs are not detected.
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

	rewrites := map[string][]byte{}
	hits := 0
	for _, f := range zr.File {
		if !(strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml")) {
			continue
		}
		data, err := common.ReadEntry(zr, f.Name)
		if err != nil {
			continue
		}
		new := data
		for _, r := range replacements {
			if r.Find == "" {
				continue
			}
			hits += strings.Count(string(new), r.Find)
			new = common.ReplaceInBytes(new, r.Find, r.Replace)
		}
		rewrites[f.Name] = new
	}

	if len(rewrites) == 0 {
		return 0, fmt.Errorf("no slide files found in %s", path)
	}

	if err := common.RewriteEntries(path, outPath, rewrites); err != nil {
		return 0, err
	}
	return hits, nil
}

// ExtractedImage describes one media file pulled from a pptx.
type ExtractedImage struct {
	Name      string `json:"name"`
	Path      string `json:"path"`
	Bytes     int64  `json:"bytes"`
	MediaType string `json:"mediaType,omitempty"`
}

// ExtractImages dumps every file under ppt/media/ into outDir. Returns metadata
// for each image written.
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
		if !strings.HasPrefix(f.Name, "ppt/media/") {
			continue
		}
		base := strings.TrimPrefix(f.Name, "ppt/media/")
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

// SlideCount returns just the number of slides — cheap version of ReadSlides.
func SlideCount(path string) (int, error) {
	zc, err := common.OpenReader(path)
	if err != nil {
		return 0, err
	}
	defer func() { _ = zc.Close() }()

	count := 0
	for _, f := range zc.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			count++
		}
	}
	return count, nil
}

// Outline is a (slide-number, title) pair used for the lightweight outline view.
type Outline struct {
	Index int    `json:"index"`
	Title string `json:"title,omitempty"`
}

// ReadOutline returns just slide titles. This is dramatically cheaper than
// ReadSlides when the user only needs the outline / table of contents.
func ReadOutline(path string) ([]Outline, error) {
	slides, err := ReadSlides(path)
	if err != nil {
		return nil, err
	}
	out := make([]Outline, 0, len(slides))
	for _, s := range slides {
		out = append(out, Outline{Index: s.Index, Title: s.Title})
	}
	return out, nil
}

// ReadMeta returns the consolidated presentation metadata.
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

	slideCount := 0
	for _, f := range zc.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
	}

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
		Company:     strings.TrimSpace(app.Company),
		Slides:      slideCount,
	}, nil
}
