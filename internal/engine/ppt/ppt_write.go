package ppt

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

// ---------------------------------------------------------------------------
// Shared helpers for rID / sldID management
// ---------------------------------------------------------------------------

// parseSldIDLst parses the <p:sldIdLst> section of presentation.xml using a
// token-based decoder. This avoids Go's xml.Unmarshal limitation where
// namespaced attributes (r:id) aren't resolved when the namespace is declared
// on a parent element.
func parseSldIDLst(presData []byte) ([]sldIDEntry, int, error) {
	dec := xml.NewDecoder(bytes.NewReader(presData))
	inSldIDLst := false
	maxSldID := 0
	var entries []sldIDEntry
	rNS := "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

	for {
		tok, tokErr := dec.Token()
		if tokErr != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "sldIdLst" {
				inSldIDLst = true
				continue
			}
			if inSldIDLst && t.Name.Local == "sldId" {
				var sldID int
				var rID string
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" && attr.Name.Space == "" {
						sldID, _ = strconv.Atoi(attr.Value)
					}
					if attr.Name.Space == rNS {
						rID = attr.Value
					}
				}
				if sldID > maxSldID {
					maxSldID = sldID
				}
				if rID != "" {
					entries = append(entries, sldIDEntry{SldID: sldID, RId: rID})
				}
			}
		case xml.EndElement:
			if t.Name.Local == "sldIdLst" {
				return entries, maxSldID, nil
			}
		}
	}
	return entries, maxSldID, nil
}

// parseSlideRelMap reads presentation.xml and its .rels file to build the real
// slide-number to rID mapping.
func parseSlideRelMap(zr *zip.Reader) (slideRelMap, error) {
	m := slideRelMap{
		slideToRId: make(map[int]string),
		rIDToSlide: make(map[string]int),
	}

	presData, err := common.ReadEntry(zr, "ppt/presentation.xml")
	if err != nil {
		return m, fmt.Errorf("reading presentation.xml: %w", err)
	}

	sldEntries, maxSldID, err := parseSldIDLst(presData)
	if err != nil {
		return m, fmt.Errorf("parsing presentation.xml: %w", err)
	}
	m.maxSldID = maxSldID

	relsData, err := common.ReadEntry(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		return m, fmt.Errorf("reading presentation.xml.rels: %w", err)
	}
	var rels xmlRels
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		return m, fmt.Errorf("parsing presentation.xml.rels: %w", err)
	}

	slideRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	for _, rel := range rels.Entries {
		n := rIDNum(rel.ID)
		if n > m.maxRIdNum {
			m.maxRIdNum = n
		}
		if rel.Type != slideRelType {
			continue
		}
		num := slideNum("ppt/"+rel.Target, "ppt/slides/slide")
		if num <= 0 {
			continue
		}
		m.slideToRId[num] = rel.ID
		m.rIDToSlide[rel.ID] = num
	}

	_ = sldEntries // mapping is done via rels
	return m, nil
}

// rIDNum extracts the numeric suffix from "rId12" → 12. Returns 0 on failure.
func rIDNum(id string) int {
	n := 0
	for _, c := range id {
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		} else if c == 'd' || c == 'D' {
			n = 0
		}
	}
	return n
}

// maxRIdNumFromBytes finds the highest rID numeric suffix in a .rels XML byte slice.
func maxRIdNumFromBytes(data []byte) int {
	var rels xmlRels
	if err := xml.Unmarshal(data, &rels); err != nil {
		return 0
	}
	max := 0
	for _, rel := range rels.Entries {
		if n := rIDNum(rel.ID); n > max {
			max = n
		}
	}
	return max
}

// xmlEscapePPT escapes XML special characters.
func xmlEscapePPT(s string) string {
	var b strings.Builder
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

func Create(path, title, author string) error {
	if _, err := os.Stat(path); err == nil {
		return fmt.Errorf("file already exists: %s", path)
	}

	titleXML := ""
	if title != "" {
		titleXML = fmt.Sprintf(`<p:sp>
          <p:nvSpPr><p:cNvPr id="1" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>%s</a:t></a:r></a:p></p:txBody>
        </p:sp>`, xmlEscapePPT(title))
	}

	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`,
		"ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
		"ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
</p:presentation>`,
		"ppt/slides/slide1.xml": fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>%s
  </p:spTree></p:cSld>
</p:sld>`, titleXML),
		"docProps/core.xml": fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>%s</dc:title>
  <dc:creator>%s</dc:creator>
</cp:coreProperties>`, xmlEscapePPT(title), xmlEscapePPT(author)),
		"docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>office-cli</Application>
  <Slides>1</Slides>
</Properties>`,
	}
	return common.WriteNewZip(path, files)
}

// makeSlideXML generates a slide XML string with optional title and bullets.
func makeSlideXML(title string, bullets []string) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>`)

	if title != "" {
		sb.WriteString(fmt.Sprintf(`
    <p:sp>
      <p:nvSpPr><p:cNvPr id="1" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>%s</a:t></a:r></a:p></p:txBody>
    </p:sp>`, xmlEscapePPT(title)))
	}

	sb.WriteString(fmt.Sprintf(`
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/>`))
	for _, b := range bullets {
		sb.WriteString(fmt.Sprintf(`<a:p><a:pPr><a:buChar char=""/></a:pPr><a:r><a:t>%s</a:t></a:r></a:p>`, xmlEscapePPT(b)))
	}
	if len(bullets) == 0 {
		sb.WriteString(`<a:p><a:r><a:t></a:t></a:r></a:p>`)
	}
	sb.WriteString(`</p:txBody>
    </p:sp>`)
	sb.WriteString(`
  </p:spTree></p:cSld>
</p:sld>`)
	return sb.String()
}

// AddSlide appends a new slide to the presentation.
// Uses parseSlideRelMap to generate safe, collision-free IDs.
func AddSlide(path, outPath, title string, bullets []string) error {
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

	// Parse real rID mapping to generate safe IDs
	relMap, err := parseSlideRelMap(zr)
	if err != nil {
		_ = srcF.Close()
		return err
	}

	// Count existing slides and find the next available slide number
	slideCount := 0
	maxFileNum := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
			n := slideNum(f.Name, "ppt/slides/slide")
			if n > maxFileNum {
				maxFileNum = n
			}
		}
	}
	newSlideNum := maxFileNum + 1
	newSlideFile := fmt.Sprintf("ppt/slides/slide%d.xml", newSlideNum)

	// Generate collision-free IDs
	newRelID := fmt.Sprintf("rID%d", relMap.maxRIdNum+1)
	newSldID := relMap.maxSldID + 1

	entries := map[string][]byte{}

	// New slide XML
	entries[newSlideFile] = []byte(makeSlideXML(title, bullets))

	// Update presentation.xml - add new sldID
	pres, err := common.ReadEntry(zr, "ppt/presentation.xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	presStr := string(pres)
	sldIDXML := fmt.Sprintf(`    <p:sldId id="%d" r:id="%s"/>`, newSldID, newRelID)
	insertPoint := strings.LastIndex(presStr, "</p:sldIdLst>")
	if insertPoint >= 0 {
		presStr = presStr[:insertPoint] + sldIDXML + "\n  " + presStr[insertPoint:]
	}
	entries["ppt/presentation.xml"] = []byte(presStr)

	// Update relationships
	rels, err := common.ReadEntry(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	relsStr := string(rels)
	newRel := fmt.Sprintf(`  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`, newRelID, newSlideNum)
	insertPoint = strings.LastIndex(relsStr, "</Relationships>")
	if insertPoint >= 0 {
		relsStr = relsStr[:insertPoint] + newRel + "\n" + relsStr[insertPoint:]
	}
	entries["ppt/_rels/presentation.xml.rels"] = []byte(relsStr)

	// Update [Content_Types].xml
	ct, err := common.ReadEntry(zr, "[Content_Types].xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	ctStr := string(ct)
	newOverride := fmt.Sprintf(`  <Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`, newSlideNum)
	insertPoint = strings.LastIndex(ctStr, "</Types>")
	if insertPoint >= 0 {
		ctStr = ctStr[:insertPoint] + newOverride + "\n" + ctStr[insertPoint:]
	}
	entries["[Content_Types].xml"] = []byte(ctStr)

	// Update app.xml slide count
	app, err := common.ReadEntry(zr, "docProps/app.xml")
	if err == nil {
		appStr := string(app)
		appStr = strings.Replace(appStr, fmt.Sprintf("<Slides>%d</Slides>", slideCount), fmt.Sprintf("<Slides>%d</Slides>", slideCount+1), 1)
		entries["docProps/app.xml"] = []byte(appStr)
	}

	_ = srcF.Close()

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, entries)
}

// SetSlideContent overwrites the text content of a specific slide.
func SetSlideContent(path, outPath string, slideNum int, title string, bullets []string) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}
	slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)

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

	// Verify slide exists
	_, err = common.ReadEntry(zr, slideFile)
	if err != nil {
		_ = srcF.Close()
		return fmt.Errorf("slide %d not found", slideNum)
	}

	newContent := makeSlideXML(title, bullets)
	out := outPath
	if out == "" {
		out = path
	}
	_ = srcF.Close()
	return common.RewriteEntries(path, out, map[string][]byte{
		slideFile: []byte(newContent),
	})
}

// SetNotes sets speaker notes for a slide.
func SetNotes(path, outPath string, slideNum int, notes string) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}

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

	notesFile := fmt.Sprintf("ppt/notesSlides/notesSlide%d.xml", slideNum)
	notesContent := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>%s</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>`, xmlEscapePPT(notes))

	entries := map[string][]byte{
		notesFile: []byte(notesContent),
	}

	// Update [Content_Types].xml
	ct, err := common.ReadEntry(zr, "[Content_Types].xml")
	if err == nil {
		ctStr := string(ct)
		override := fmt.Sprintf(`  <Override PartName="/ppt/notesSlides/notesSlide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`, slideNum)
		if !strings.Contains(ctStr, fmt.Sprintf("notesSlide%d.xml", slideNum)) {
			insertPoint := strings.LastIndex(ctStr, "</Types>")
			if insertPoint >= 0 {
				ctStr = ctStr[:insertPoint] + override + "\n" + ctStr[insertPoint:]
			}
		}
		entries["[Content_Types].xml"] = []byte(ctStr)
	}

	// Register the notesSlide relationship from the slide's .rels file.
	// Without this, PowerPoint will not show the notes panel for this slide.
	slideRelsFile := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
	slideRelsData, err := common.ReadEntry(zr, slideRelsFile)
	notesSlideTarget := fmt.Sprintf("../notesSlides/notesSlide%d.xml", slideNum)
	notesRelType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
	if err == nil {
		// Existing .rels file — check if relationship already exists
		slideRelsStr := string(slideRelsData)
		if !strings.Contains(slideRelsStr, notesSlideTarget) {
			nextNum := maxRIdNumFromBytes(slideRelsData) + 1
			newRel := fmt.Sprintf(`  <Relationship Id="rID%d" Type="%s" Target="%s"/>`, nextNum, notesRelType, notesSlideTarget)
			insertPoint := strings.LastIndex(slideRelsStr, "</Relationships>")
			if insertPoint >= 0 {
				slideRelsStr = slideRelsStr[:insertPoint] + newRel + "\n" + slideRelsStr[insertPoint:]
			}
			entries[slideRelsFile] = []byte(slideRelsStr)
		}
	} else {
		// No .rels file exists yet — create one
		newRels := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="%s" Target="%s"/>
</Relationships>`, notesRelType, notesSlideTarget)
		entries[slideRelsFile] = []byte(newRels)
	}

	_ = srcF.Close()

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, entries)
}

// DeleteSlide removes a slide by number.
func DeleteSlide(path, outPath string, slideNum int) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}

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

	slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)

	// Read the files we need to modify
	_, err = common.ReadEntry(zr, slideFile)
	if err != nil {
		_ = srcF.Close()
		return fmt.Errorf("slide %d not found", slideNum)
	}

	// Parse the real slide ↔ rID mapping
	relMap, err := parseSlideRelMap(zr)
	if err != nil {
		_ = srcF.Close()
		return err
	}
	relID, ok := relMap.slideToRId[slideNum]
	if !ok {
		_ = srcF.Close()
		return fmt.Errorf("no relationship found for slide %d", slideNum)
	}

	// Remove the slide from presentation.xml
	pres, err := common.ReadEntry(zr, "ppt/presentation.xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	presStr := string(pres)
	sldIDPattern := fmt.Sprintf(`r:id="%s"`, relID)
	lines := strings.Split(presStr, "\n")
	var newLines []string
	for _, line := range lines {
		if strings.Contains(line, sldIDPattern) {
			continue // skip this slide entry
		}
		newLines = append(newLines, line)
	}

	// Remove relationship
	rels, err := common.ReadEntry(zr, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	relsStr := string(rels)
	relLines := strings.Split(relsStr, "\n")
	var newRelLines []string
	for _, line := range relLines {
		if strings.Contains(line, fmt.Sprintf(`Id="%s"`, relID)) {
			continue
		}
		newRelLines = append(newRelLines, line)
	}

	// Remove from [Content_Types].xml
	ct, err := common.ReadEntry(zr, "[Content_Types].xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	ctStr := string(ct)
	ctLines := strings.Split(ctStr, "\n")
	var newCtLines []string
	for _, line := range ctLines {
		if strings.Contains(line, fmt.Sprintf("slide%d.xml", slideNum)) {
			continue
		}
		newCtLines = append(newCtLines, line)
	}

	// Update app.xml slide count
	app, _ := common.ReadEntry(zr, "docProps/app.xml")
	appStr := string(app)
	currentCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			currentCount++
		}
	}
	appStr = strings.Replace(appStr, fmt.Sprintf("<Slides>%d</Slides>", currentCount), fmt.Sprintf("<Slides>%d</Slides>", currentCount-1), 1)

	_ = srcF.Close()

	out := outPath
	if out == "" {
		out = path
	}

	// Rewrite everything except the deleted slide
	entries := map[string][]byte{
		"ppt/presentation.xml":            []byte(strings.Join(newLines, "\n")),
		"ppt/_rels/presentation.xml.rels": []byte(strings.Join(newRelLines, "\n")),
		"[Content_Types].xml":             []byte(strings.Join(newCtLines, "\n")),
		"docProps/app.xml":                []byte(appStr),
	}

	// Read source into memory to allow safe in-place overwriting.
	srcData, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	zr2, err := zip.NewReader(bytes.NewReader(srcData), int64(len(srcData)))
	if err != nil {
		return err
	}

	outF, err := os.Create(out)
	if err != nil {
		return err
	}

	zw := zip.NewWriter(outF)
	for _, f := range zr2.File {
		if f.Name == slideFile {
			continue // skip deleted slide
		}
		w, err := zw.CreateHeader(&zip.FileHeader{Name: f.Name, Method: f.Method})
		if err != nil {
			_ = zw.Close()
			_ = outF.Close()
			return err
		}
		if newBytes, ok := entries[f.Name]; ok {
			if _, err := w.Write(newBytes); err != nil {
				_ = zw.Close()
				_ = outF.Close()
				return err
			}
			continue
		}
		rc, err := f.Open()
		if err != nil {
			_ = zw.Close()
			_ = outF.Close()
			return err
		}
		if _, err := io.Copy(w, rc); err != nil {
			_ = rc.Close()
			_ = zw.Close()
			_ = outF.Close()
			return err
		}
		_ = rc.Close()
	}
	if err := zw.Close(); err != nil {
		_ = outF.Close()
		return err
	}
	return outF.Close()
}

// ReorderSlides reorders slides according to the given order (e.g. "3,1,2").
func ReorderSlides(path, outPath, orderStr string) error {
	parts := strings.Split(orderStr, ",")
	if len(parts) == 0 {
		return errors.New("empty order")
	}
	var order []int
	for _, p := range parts {
		n, err := strconv.Atoi(strings.TrimSpace(p))
		if err != nil || n < 1 {
			return fmt.Errorf("invalid slide number in order: %q", strings.TrimSpace(p))
		}
		order = append(order, n)
	}

	// For reorder, we just need to rewrite presentation.xml with the new order
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

	pres, err := common.ReadEntry(zr, "ppt/presentation.xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}

	// Parse the real slide ↔ rID mapping
	relMap, err := parseSlideRelMap(zr)
	if err != nil {
		_ = srcF.Close()
		return err
	}

	// Build new sldIdLst with the requested order
	// Each slide keeps its original sldID to avoid breaking internal references
	var sldIDLst strings.Builder
	sldIDLst.WriteString("  <p:sldIdLst>\n")
	nextSldID := relMap.maxSldID + 1
	for _, slideNum := range order {
		rid, ok := relMap.slideToRId[slideNum]
		if !ok {
			_ = srcF.Close()
			return fmt.Errorf("no relationship found for slide %d", slideNum)
		}
		sldIDLst.WriteString(fmt.Sprintf("    <p:sldId id=\"%d\" r:id=\"%s\"/>\n", nextSldID, rid))
		nextSldID++
	}
	sldIDLst.WriteString("  </p:sldIdLst>")

	presStr := string(pres)
	// Replace the entire sldIdLst
	startIdx := strings.Index(presStr, "<p:sldIdLst>")
	endIdx := strings.Index(presStr, "</p:sldIdLst>")
	if startIdx >= 0 && endIdx >= 0 {
		presStr = presStr[:startIdx] + sldIDLst.String() + presStr[endIdx+len("</p:sldIdLst>"):]
	}

	_ = srcF.Close()

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		"ppt/presentation.xml": []byte(presStr),
	})
}

// AddImage inserts a real image into a specific slide using <p:pic> with <a:blip>.
// The image is embedded in ppt/media/ and referenced via a slide relationship.
// widthEMU and heightEMU control the rendered size (0 = default 10"×7.5" full-slide).
func AddImage(path, outPath string, slideNum int, imagePath string, widthEMU, heightEMU int) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}

	imgData, err := os.ReadFile(imagePath)
	if err != nil {
		return fmt.Errorf("reading image: %w", err)
	}

	ext := strings.ToLower(filepath.Ext(imagePath))
	ct := common.GuessMediaType(filepath.Base(imagePath))
	if ct == "" {
		ct = "image/png"
	}

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

	// Count existing media to generate unique names
	mediaCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}

	mediaName := fmt.Sprintf("image%d%s", mediaCount+1, ext)
	slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	relsFile := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)

	// Find the next available rID for this slide's relationships
	var nextRIdNum int
	relsData, err := common.ReadEntry(zr, relsFile)
	relsStr := ""
	if err == nil {
		relsStr = string(relsData)
		nextRIdNum = maxRIdNumFromBytes(relsData) + 1
	} else {
		relsStr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
		nextRIdNum = 1
	}
	rID := fmt.Sprintf("rID%d", nextRIdNum)

	// Read existing slide
	slideData, err := common.ReadEntry(zr, slideFile)
	if err != nil {
		_ = srcF.Close()
		return fmt.Errorf("slide %d not found", slideNum)
	}

	// Default to full slide (10"×7.5" in EMU: 1 inch = 914400 EMU)
	if widthEMU <= 0 {
		widthEMU = 9144000
	}
	if heightEMU <= 0 {
		heightEMU = 6858000
	}

	// Build a proper <p:pic> element with <a:blip> for real image rendering
	picID := mediaCount + 10
	imgShape := fmt.Sprintf(`    <p:pic>
      <p:nvPicPr><p:cNvPr id="%d" name="Picture %d"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>`, picID, mediaCount+1, rID, widthEMU, heightEMU)

	slideStr := string(slideData)
	insertPoint := strings.LastIndex(slideStr, "</p:spTree>")
	if insertPoint < 0 {
		_ = srcF.Close()
		return fmt.Errorf("invalid slide XML: missing </p:spTree>")
	}
	slideStr = slideStr[:insertPoint] + imgShape + "\n" + slideStr[insertPoint:]

	entries := map[string][]byte{
		"ppt/media/" + mediaName: imgData,
		slideFile:                []byte(slideStr),
	}

	// Add image relationship to slide .rels
	newRel := fmt.Sprintf(`  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/>`, rID, mediaName)
	insertPoint = strings.LastIndex(relsStr, "</Relationships>")
	if insertPoint >= 0 {
		relsStr = relsStr[:insertPoint] + newRel + "\n" + relsStr[insertPoint:]
	}
	entries[relsFile] = []byte(relsStr)

	// Update [Content_Types].xml
	ctData, _ := common.ReadEntry(zr, "[Content_Types].xml")
	if ctData != nil {
		ctStr := string(ctData)
		extNoDot := strings.TrimPrefix(ext, ".")
		if !strings.Contains(ctStr, fmt.Sprintf(`Extension="%s"`, extNoDot)) {
			newCT := fmt.Sprintf(`  <Default Extension="%s" ContentType="%s"/>`, extNoDot, ct)
			idx := strings.LastIndex(ctStr, "</Types>")
			if idx >= 0 {
				ctStr = ctStr[:idx] + newCT + "\n" + ctStr[idx:]
			}
		}
		entries["[Content_Types].xml"] = []byte(ctStr)
	}

	_ = srcF.Close()

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, entries)
}

// ---------------------------------------------------------------------------
// Build — one-shot deck creation from a JSON spec
// ---------------------------------------------------------------------------

// SlideSpec describes a single slide in a BuildSpec.
type SlideSpec struct {
	Title   string   `json:"title,omitempty"`
	Bullets []string `json:"bullets,omitempty"`
	Notes   string   `json:"notes,omitempty"`
	Image   string   `json:"image,omitempty"`  // path to image file to embed
	Width   int      `json:"width,omitempty"`  // image width in EMU (0 = full slide width)
	Height  int      `json:"height,omitempty"` // image height in EMU (0 = full slide height)
}

// BuildSpec is the JSON payload for the Build function.
type BuildSpec struct {
	Title  string      `json:"title,omitempty"`
	Author string      `json:"author,omitempty"`
	Slides []SlideSpec `json:"slides"`
}

// Build creates a complete .pptx from a JSON spec in one call.
// It internally calls Create, SetSlideContent (for the first slide), AddSlide
// (for subsequent slides), SetNotes (when notes are present), and AddImage
// (when an image path is specified per slide).
func Build(path string, spec BuildSpec) error {
	if len(spec.Slides) == 0 {
		return errors.New("slides array must not be empty")
	}

	// 1. Create a minimal presentation (produces 1 empty slide)
	if err := Create(path, spec.Title, spec.Author); err != nil {
		return fmt.Errorf("creating presentation: %w", err)
	}

	// 2. Replace the first slide's content
	first := spec.Slides[0]
	if err := SetSlideContent(path, path, 1, first.Title, first.Bullets); err != nil {
		return fmt.Errorf("setting slide 1 content: %w", err)
	}
	if first.Notes != "" {
		if err := SetNotes(path, path, 1, first.Notes); err != nil {
			return fmt.Errorf("setting slide 1 notes: %w", err)
		}
	}
	if first.Image != "" {
		if err := AddImage(path, path, 1, first.Image, first.Width, first.Height); err != nil {
			return fmt.Errorf("adding image to slide 1: %w", err)
		}
	}

	// 3. Add remaining slides
	for i := 1; i < len(spec.Slides); i++ {
		s := spec.Slides[i]
		slideNum := i + 1
		if err := AddSlide(path, path, s.Title, s.Bullets); err != nil {
			return fmt.Errorf("adding slide %d: %w", slideNum, err)
		}
		if s.Notes != "" {
			if err := SetNotes(path, path, slideNum, s.Notes); err != nil {
				return fmt.Errorf("setting slide %d notes: %w", slideNum, err)
			}
		}
		if s.Image != "" {
			if err := AddImage(path, path, slideNum, s.Image, s.Width, s.Height); err != nil {
				return fmt.Errorf("adding image to slide %d: %w", slideNum, err)
			}
		}
	}

	return nil
}

// BuildFromTemplate creates a deck using an existing .pptx as a template.
// The template's slide master, layouts, themes, and fonts are preserved.
// All existing slides are removed; new slides are created from the spec.
func BuildFromTemplate(templatePath, outPath string, spec BuildSpec) error {
	if len(spec.Slides) == 0 {
		return errors.New("slides array must not be empty")
	}
	if _, err := os.Stat(templatePath); err != nil {
		return fmt.Errorf("template not found: %w", err)
	}

	// 1. Copy template to output path
	if err := copyFile(templatePath, outPath); err != nil {
		return fmt.Errorf("copying template: %w", err)
	}

	// 2. Count existing slides in the template
	srcF, err := os.Open(outPath)
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
	totalSlides := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			totalSlides++
		}
	}
	_ = srcF.Close()

	// 3. Delete all slides except the first (from end to avoid renumbering issues)
	for i := totalSlides; i > 1; i-- {
		if err := DeleteSlide(outPath, outPath, i); err != nil {
			return fmt.Errorf("removing template slide %d: %w", i, err)
		}
	}

	// 4. Update slide 1 with first spec entry
	first := spec.Slides[0]
	if err := SetSlideContent(outPath, outPath, 1, first.Title, first.Bullets); err != nil {
		return fmt.Errorf("setting slide 1 content: %w", err)
	}
	if first.Notes != "" {
		if err := SetNotes(outPath, outPath, 1, first.Notes); err != nil {
			return fmt.Errorf("setting slide 1 notes: %w", err)
		}
	}
	if first.Image != "" {
		if err := AddImage(outPath, outPath, 1, first.Image, first.Width, first.Height); err != nil {
			return fmt.Errorf("adding image to slide 1: %w", err)
		}
	}

	// 5. Add remaining slides
	for i := 1; i < len(spec.Slides); i++ {
		s := spec.Slides[i]
		slideNum := i + 1
		if err := AddSlide(outPath, outPath, s.Title, s.Bullets); err != nil {
			return fmt.Errorf("adding slide %d: %w", slideNum, err)
		}
		if s.Notes != "" {
			if err := SetNotes(outPath, outPath, slideNum, s.Notes); err != nil {
				return fmt.Errorf("setting slide %d notes: %w", slideNum, err)
			}
		}
		if s.Image != "" {
			if err := AddImage(outPath, outPath, slideNum, s.Image, s.Width, s.Height); err != nil {
				return fmt.Errorf("adding image to slide %d: %w", slideNum, err)
			}
		}
	}

	// 6. Update metadata (title/author) in core.xml
	if spec.Title != "" || spec.Author != "" {
		if err := updateCoreProps(outPath, spec.Title, spec.Author); err != nil {
			return fmt.Errorf("updating metadata: %w", err)
		}
	}

	return nil
}

// copyFile copies a file from src to dst.
func copyFile(src, dst string) error {
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, data, 0644)
}

// updateCoreProps rewrites docProps/core.xml with new title and/or author.
func updateCoreProps(path, title, author string) error {
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

	coreData, err := common.ReadEntry(zr, "docProps/core.xml")
	if err != nil {
		_ = srcF.Close()
		return err
	}
	_ = srcF.Close()

	coreStr := string(coreData)
	if title != "" {
		coreStr = replaceXMLTag(coreStr, "dc:title", title)
	}
	if author != "" {
		coreStr = replaceXMLTag(coreStr, "dc:creator", author)
	}

	return common.RewriteEntries(path, path, map[string][]byte{
		"docProps/core.xml": []byte(coreStr),
	})
}

// replaceXMLTag replaces the content of an XML tag (e.g. <dc:title>old</dc:title> → <dc:title>new</dc:title>).
func replaceXMLTag(xml, tag, value string) string {
	open := "<" + tag + ">"
	close := "</" + tag + ">"
	start := strings.Index(xml, open)
	if start < 0 {
		return xml
	}
	end := strings.Index(xml[start:], close)
	if end < 0 {
		return xml
	}
	end += start + len(close)
	return xml[:start] + open + xmlEscapePPT(value) + close + xml[end:]
}

// ---------------------------------------------------------------------------
// ReadSlideLayout — shape tree reader for layout awareness
