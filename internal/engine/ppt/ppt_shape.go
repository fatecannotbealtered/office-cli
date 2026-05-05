package ppt

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
)

func ReadSlideLayout(path string, slideNum int) ([]ShapeInfo, error) {
	srcF, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	stat, err := srcF.Stat()
	if err != nil {
		_ = srcF.Close()
		return nil, err
	}
	zr, err := zip.NewReader(srcF, stat.Size())
	if err != nil {
		_ = srcF.Close()
		return nil, err
	}
	defer func() { _ = srcF.Close() }()

	slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	data, err := common.ReadEntry(zr, slideFile)
	if err != nil {
		return nil, fmt.Errorf("slide %d not found", slideNum)
	}

	return parseShapeTree(data)
}

// parseShapeTree extracts shape info from slide XML using string-based extraction.
// This avoids the complexity of depth tracking with xml.Decoder.
func parseShapeTree(data []byte) ([]ShapeInfo, error) {
	xmlStr := string(data)
	aNS := "http://schemas.openxmlformats.org/drawingml/2006/main"
	pNS := "http://schemas.openxmlformats.org/presentationml/2006/main"

	// Find spTree boundaries
	spTreeStart := strings.Index(xmlStr, "<p:spTree")
	if spTreeStart < 0 {
		spTreeStart = strings.Index(xmlStr, "<spTree")
	}
	if spTreeStart < 0 {
		return nil, nil
	}

	var shapes []ShapeInfo
	remaining := xmlStr[spTreeStart:]

	for {
		// Find next <p:sp or <sp (shape element)
		spIdx := findShapeTag(remaining, "sp")
		if spIdx < 0 {
			break
		}
		startOffset := spIdx
		endOffset := findMatchingEnd(remaining, startOffset, "sp")
		if endOffset < 0 {
			break
		}
		shapeXML := remaining[startOffset:endOffset]
		si := parseSingleShape(shapeXML, aNS, pNS, "sp", len(shapes))
		shapes = append(shapes, si)
		remaining = remaining[endOffset:]
	}

	// Also look for <p:pic> elements
	remaining = xmlStr[spTreeStart:]
	for {
		picIdx := findShapeTag(remaining, "pic")
		if picIdx < 0 {
			break
		}
		startOffset := picIdx
		endOffset := findMatchingEnd(remaining, startOffset, "pic")
		if endOffset < 0 {
			break
		}
		shapeXML := remaining[startOffset:endOffset]
		si := parseSingleShape(shapeXML, aNS, pNS, "pic", len(shapes))
		shapes = append(shapes, si)
		remaining = remaining[endOffset:]
	}

	return shapes, nil
}

// findShapeTag finds the byte offset of the next <p:TAG or <TAG in xmlStr.
// It loops through all occurrences to skip false matches like <p:spTree when looking for <p:sp.
func findShapeTag(xmlStr, localName string) int {
	// Try prefixed form first: <p:sp or <p:pic
	prefixed := "<p:" + localName
	for i := 0; i < len(xmlStr); i++ {
		j := strings.Index(xmlStr[i:], prefixed)
		if j < 0 {
			break
		}
		j += i
		afterIdx := j + len(prefixed)
		if afterIdx < len(xmlStr) {
			ch := xmlStr[afterIdx]
			if ch == '>' || ch == ' ' || ch == '/' {
				return j
			}
		}
		i = afterIdx
	}
	// Try unprefixed: <sp (but not <spTree, <spPr, etc.)
	unprefixed := "<" + localName
	for i := 0; i < len(xmlStr); i++ {
		j := strings.Index(xmlStr[i:], unprefixed)
		if j < 0 {
			break
		}
		j += i
		afterIdx := j + len(unprefixed)
		if afterIdx < len(xmlStr) {
			ch := xmlStr[afterIdx]
			if ch == '>' || ch == ' ' || ch == '/' {
				return j
			}
		}
		i = afterIdx
	}
	return -1
}

// findMatchingEnd returns the byte offset AFTER the matching closing tag for an element
// starting at startPos in xmlStr.
func findMatchingEnd(xmlStr string, startPos int, localName string) int {
	depth := 0
	prefixed := "<p:" + localName
	unprefixed := "<" + localName
	closePrefixed := "</p:" + localName + ">"
	closeUnprefixed := "</" + localName + ">"

	i := startPos
	for i < len(xmlStr) {
		if xmlStr[i] != '<' {
			i++
			continue
		}
		// Check for start tag
		isStart := false
		if strings.HasPrefix(xmlStr[i:], prefixed) {
			after := i + len(prefixed)
			if after < len(xmlStr) && (xmlStr[after] == '>' || xmlStr[after] == ' ' || xmlStr[after] == '/') {
				isStart = true
			}
		} else if strings.HasPrefix(xmlStr[i:], unprefixed) {
			after := i + len(unprefixed)
			if after < len(xmlStr) && (xmlStr[after] == '>' || xmlStr[after] == ' ' || xmlStr[after] == '/') {
				isStart = true
			}
		}
		if isStart {
			// Check for self-closing
			closeAngle := strings.Index(xmlStr[i:], ">")
			if closeAngle >= 0 && xmlStr[i+closeAngle-1] == '/' {
				// Self-closing, don't increment depth
			} else {
				depth++
			}
			i++
			continue
		}
		// Check for end tag
		if strings.HasPrefix(xmlStr[i:], closePrefixed) || strings.HasPrefix(xmlStr[i:], closeUnprefixed) {
			depth--
			if depth == 0 {
				endTagLen := len(closePrefixed)
				if strings.HasPrefix(xmlStr[i:], closeUnprefixed) {
					endTagLen = len(closeUnprefixed)
				}
				return i + endTagLen
			}
		}
		i++
	}
	return -1
}

// parseSingleShape parses a single shape's XML (extracted as a string) into a ShapeInfo.
func parseSingleShape(shapeXML string, aNS, pNS, shapeType string, idx int) ShapeInfo {
	si := ShapeInfo{Index: idx, Type: shapeType}

	// Extract name from <cNvPr ... name="..." .../> or <p:cNvPr ... name="..." .../>
	if m := regexp.MustCompile(`<(?:\w+:)?cNvPr[^>]+name="([^"]*)"`).FindStringSubmatch(shapeXML); len(m) > 1 {
		si.Name = m[1]
	}

	// Extract placeholder type from <ph type="..."/> or <p:ph type="..."/>
	if m := regexp.MustCompile(`<(?:\w+:)?ph[^>]+type="([^"]*)"`).FindStringSubmatch(shapeXML); len(m) > 1 {
		si.Ph = m[1]
	}

	// Extract position/size from <a:off x="..." y="..."/> and <a:ext cx="..." cy="..."/>
	if m := regexp.MustCompile(`<a:off[^>]+x="(\d+)"[^>]+y="(\d+)"`).FindStringSubmatch(shapeXML); len(m) > 2 {
		si.X, _ = strconv.Atoi(m[1])
		si.Y, _ = strconv.Atoi(m[2])
	}
	if m := regexp.MustCompile(`<a:ext[^>]+cx="(\d+)"[^>]+cy="(\d+)"`).FindStringSubmatch(shapeXML); len(m) > 2 {
		si.W, _ = strconv.Atoi(m[1])
		si.H, _ = strconv.Atoi(m[2])
	}

	// Extract text paragraphs from <p:txBody> or <txBody>
	// Find txBody content
	txBodyStart := strings.Index(shapeXML, "<p:txBody")
	if txBodyStart < 0 {
		txBodyStart = strings.Index(shapeXML, "<txBody")
	}
	if txBodyStart >= 0 {
		txBodyEnd := strings.LastIndex(shapeXML, "</p:txBody>")
		if txBodyEnd < 0 {
			txBodyEnd = strings.LastIndex(shapeXML, "</txBody>")
		}
		if txBodyEnd > txBodyStart {
			txBodyXML := shapeXML[txBodyStart : txBodyEnd+len("</txBody>")]
			if strings.Contains(txBodyXML, "</p:txBody>") {
				txBodyXML = shapeXML[txBodyStart : txBodyEnd+len("</p:txBody>")]
			}
			si.Paragraphs, si.Text = parseTxBodyString(txBodyXML, aNS)
		}
	}

	return si
}

// parseTxBodyString parses paragraphs from a txBody XML string.
func parseTxBodyString(txBodyXML string, aNS string) ([]ParagraphInfo, string) {
	var paragraphs []ParagraphInfo
	var allText []string

	// Find each <a:p>...</a:p> or <p>...</p>
	remaining := txBodyXML
	for {
		pStart := strings.Index(remaining, "<a:p>")
		if pStart < 0 {
			pStart = strings.Index(remaining, "<a:p ")
		}
		if pStart < 0 {
			pStart = strings.Index(remaining, "<p>")
		}
		if pStart < 0 {
			break
		}
		// Find matching </a:p> or </p>
		pEnd := strings.Index(remaining[pStart:], "</a:p>")
		endLen := len("</a:p>")
		if pEnd < 0 {
			pEnd = strings.Index(remaining[pStart:], "</p>")
			endLen = len("</p>")
		}
		if pEnd < 0 {
			break
		}
		pEnd += pStart + endLen
		paraXML := remaining[pStart:pEnd]
		remaining = remaining[pEnd:]

		pi := parseParagraphString(paraXML, aNS)
		if len(pi.Runs) > 0 {
			paragraphs = append(paragraphs, pi)
			var sb strings.Builder
			for _, r := range pi.Runs {
				sb.WriteString(r.Text)
			}
			if s := sb.String(); s != "" {
				allText = append(allText, s)
			}
		}
	}

	return paragraphs, strings.Join(allText, "\n")
}

// parseParagraphString parses a single paragraph from its XML string.
func parseParagraphString(paraXML string, aNS string) ParagraphInfo {
	var pi ParagraphInfo

	// Extract alignment from <a:pPr ... algn="..."/>
	if m := regexp.MustCompile(`<a:pPr[^>]+algn="([^"]*)"`).FindStringSubmatch(paraXML); len(m) > 1 {
		pi.Align = m[1]
	}

	// Find each <a:r>...</a:r>
	remaining := paraXML
	for {
		rStart := strings.Index(remaining, "<a:r>")
		if rStart < 0 {
			rStart = strings.Index(remaining, "<a:r ")
		}
		if rStart < 0 {
			break
		}
		rEnd := strings.Index(remaining[rStart:], "</a:r>")
		if rEnd < 0 {
			break
		}
		rEnd += rStart + len("</a:r>")
		runXML := remaining[rStart:rEnd]
		remaining = remaining[rEnd:]

		ri := parseRunString(runXML)
		pi.Runs = append(pi.Runs, ri)
	}

	return pi
}

// parseRunString parses a single text run from its XML string.
func parseRunString(runXML string) RunInfo {
	var ri RunInfo

	// Extract text from <a:t>...</a:t>
	if m := regexp.MustCompile(`<a:t>([^<]*)</a:t>`).FindStringSubmatch(runXML); len(m) > 1 {
		ri.Text = xmlUnescape(m[1])
	}

	// Extract rPr attributes: <a:rPr ... sz="..." b="1" i="1" u="sng" ...>
	rprStart := strings.Index(runXML, "<a:rPr")
	if rprStart < 0 {
		return ri
	}
	rprEnd := strings.Index(runXML[rprStart:], ">")
	if rprEnd < 0 {
		return ri
	}
	rprEnd += rprStart
	rprTag := runXML[rprStart : rprEnd+1]

	if m := regexp.MustCompile(`sz="(\d+)"`).FindStringSubmatch(rprTag); len(m) > 1 {
		ri.FontSize, _ = strconv.Atoi(m[1])
	}
	if strings.Contains(rprTag, `b="1"`) || strings.Contains(rprTag, `b="true"`) {
		ri.Bold = true
	}
	if strings.Contains(rprTag, `i="1"`) || strings.Contains(rprTag, `i="true"`) {
		ri.Italic = true
	}
	if strings.Contains(rprTag, `u="sng"`) || strings.Contains(rprTag, `u="dbl"`) {
		ri.Underline = true
	}

	// Extract color from <a:solidFill><a:srgbClr val="..."/></a:solidFill>
	if m := regexp.MustCompile(`<a:srgbClr[^>]+val="([0-9A-Fa-f]{6})"`).FindStringSubmatch(runXML); len(m) > 1 {
		ri.Color = strings.ToUpper(m[1])
	}

	return ri
}

// xmlUnescape handles basic XML entities.
func xmlUnescape(s string) string {
	s = strings.ReplaceAll(s, "&amp;", "&")
	s = strings.ReplaceAll(s, "&lt;", "<")
	s = strings.ReplaceAll(s, "&gt;", ">")
	s = strings.ReplaceAll(s, "&quot;", "\"")
	s = strings.ReplaceAll(s, "&apos;", "'")
	return s
}

// ---------------------------------------------------------------------------
// SetShapeStyle — modify text styling within a specific shape
// ---------------------------------------------------------------------------

// SetShapeStyle modifies text properties (font size, bold, italic, color, alignment)
// within the shape at shapeIndex (0-based) on slide slideNum (1-based).
func SetShapeStyle(path, outPath string, slideNum, shapeIndex int, opts StyleOptions) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}
	if shapeIndex < 0 {
		return fmt.Errorf("shape index must be >= 0, got %d", shapeIndex)
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
	data, err := common.ReadEntry(zr, slideFile)
	if err != nil {
		_ = srcF.Close()
		return fmt.Errorf("slide %d not found", slideNum)
	}
	_ = srcF.Close()

	xmlStr := string(data)

	// Find the byte boundaries of the Nth <p:sp> element using InputOffset tracking.
	shapeStart, shapeEnd, err := findShapeRange(xmlStr, shapeIndex)
	if err != nil {
		return err
	}

	shapeXML := xmlStr[shapeStart:shapeEnd]

	// Modify <a:rPr> tags within the shape
	shapeXML = styleRuns(shapeXML, opts)

	// Modify <a:pPr> alignment
	if opts.Align != nil {
		shapeXML = styleAlign(shapeXML, *opts.Align)
	}

	newXML := xmlStr[:shapeStart] + shapeXML + xmlStr[shapeEnd:]

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		slideFile: []byte(newXML),
	})
}

// findShapeRange returns the start/end byte offsets of the shapeIndex-th <p:sp> in xmlStr.
func findShapeRange(xmlStr string, shapeIndex int) (int, int, error) {
	pNS := "http://schemas.openxmlformats.org/presentationml/2006/main"
	dec := xml.NewDecoder(strings.NewReader(xmlStr))
	inSpTree := false
	shapeIdx := 0

	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		t, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}

		if t.Name.Space == pNS && t.Name.Local == "spTree" {
			inSpTree = true
			continue
		}
		if !inSpTree {
			continue
		}
		if t.Name.Space != pNS || t.Name.Local != "sp" {
			continue
		}

		if shapeIdx == shapeIndex {
			// We're at the start of the target <sp>.
			// InputOffset() gives us bytes consumed so far; the <sp> start is at the
			// position just before the current token. We'll use a trick: find "<sp"
			// or "<p:sp" in the raw string near the current offset.
			offset := int(dec.InputOffset())
			start := findTagStart(xmlStr, offset)
			// Find matching </p:sp> or </sp> by counting depth
			end := findMatchingEndXML(xmlStr, start, "sp", pNS)
			if end < 0 {
				return 0, 0, fmt.Errorf("could not find end of shape %d", shapeIndex)
			}
			return start, end, nil
		}
		shapeIdx++
	}
	return 0, 0, fmt.Errorf("shape index %d out of range (found %d shapes)", shapeIndex, shapeIdx)
}

// findTagStart walks backwards from the approximate offset to find the '<' of the tag.
func findTagStart(xmlStr string, approxOffset int) int {
	for i := approxOffset; i >= 0; i-- {
		if xmlStr[i] == '<' {
			return i
		}
	}
	return 0
}

// findMatchingEndXML finds the byte offset after the matching closing tag for an element
// starting at startPos, using proper XML decoding.
func findMatchingEndXML(xmlStr string, startPos int, localName, space string) int {
	dec := xml.NewDecoder(strings.NewReader(xmlStr[startPos:]))
	// Skip past the opening token
	tok, err := dec.Token()
	if err != nil {
		return -1
	}
	if _, ok := tok.(xml.StartElement); !ok {
		return -1
	}

	depth := 1
	for depth > 0 {
		tok, err = dec.Token()
		if err != nil {
			break
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return startPos + int(dec.InputOffset())
}

// styleRuns modifies <a:rPr> attributes within the given XML fragment.
func styleRuns(xml string, opts StyleOptions) string {
	// String-level replacement of all <a:rPr> elements.
	// We find each <a:rPr, determine if it's self-closing or has children,
	// extract the full span, build a replacement, and advance past it.
	result := xml
	searchFrom := 0

	for searchFrom < len(result) {
		idx := strings.Index(result[searchFrom:], "<a:rPr")
		if idx < 0 {
			break
		}
		idx += searchFrom

		// Find the end of the opening tag (either /> or >)
		tagEnd := idx + 6 // past "<a:rPr"
		selfClose := false
		for tagEnd < len(result) {
			if result[tagEnd] == '/' && tagEnd+1 < len(result) && result[tagEnd+1] == '>' {
				selfClose = true
				tagEnd += 2 // past "/>"
				break
			}
			if result[tagEnd] == '>' {
				tagEnd++ // past ">"
				break
			}
			tagEnd++
		}

		// Determine the full element span (opening tag through closing tag or self-close)
		elemEnd := tagEnd
		if !selfClose {
			// Find </a:rPr> to get the full element
			closeIdx := strings.Index(result[tagEnd:], "</a:rPr>")
			if closeIdx >= 0 {
				elemEnd = tagEnd + closeIdx + len("</a:rPr>")
			}
		}

		// Parse attributes from the opening tag
		tagContent := result[idx+6 : tagEnd-1] // between <a:rPr and >
		if selfClose {
			tagContent = result[idx+6 : tagEnd-2]
		}
		attrs := parseAttrString(tagContent)

		// Apply style options
		if opts.FontSize != nil {
			attrs["sz"] = fmt.Sprintf("%d", *opts.FontSize)
		}
		if opts.Bold != nil {
			if *opts.Bold {
				attrs["b"] = "1"
			} else {
				delete(attrs, "b")
			}
		}
		if opts.Italic != nil {
			if *opts.Italic {
				attrs["i"] = "1"
			} else {
				delete(attrs, "i")
			}
		}
		if opts.Underline != nil {
			if *opts.Underline {
				attrs["u"] = "sng"
			} else {
				delete(attrs, "u")
			}
		}
		attrs["lang"] = "en-US"

		// Build the new opening tag
		newTag := "<a:rPr"
		for k, v := range attrs {
			newTag += fmt.Sprintf(` %s="%s"`, k, v)
		}

		// Build the full replacement element
		var newElem string
		needColor := opts.Color != nil && *opts.Color != ""
		colorFill := "<a:solidFill><a:srgbClr val=\"" + *opts.Color + "\"/></a:solidFill>"

		if needColor {
			// Always produce a non-self-closing element with solidFill child
			newElem = newTag + ">" + colorFill + "</a:rPr>"
		} else if selfClose {
			newElem = newTag + "/>"
		} else {
			// Keep existing children (e.g. solidFill already present)
			children := result[tagEnd : elemEnd-len("</a:rPr>")]
			newElem = newTag + ">" + children + "</a:rPr>"
		}

		result = result[:idx] + newElem + result[elemEnd:]
		// Advance past the replacement to avoid re-matching
		searchFrom = idx + len(newElem)
	}

	return result
}

// styleAlign modifies <a:pPr algn="..."> attributes within the given XML fragment.
// If <a:pPr> doesn't exist, it inserts one after each <a:p>.
func styleAlign(xmlStr string, align string) string {
	// Map friendly names to OOXML values
	ooxmlAlign := map[string]string{
		"left":    "l",
		"right":   "r",
		"center":  "ctr",
		"justify": "just",
		"l":       "l",
		"r":       "r",
		"ctr":     "ctr",
		"just":    "just",
	}
	alVal, ok := ooxmlAlign[strings.ToLower(align)]
	if !ok {
		return xmlStr
	}

	result := xmlStr

	// If <a:pPr> doesn't exist, insert one after each <a:p> opening tag
	if !strings.Contains(result, "<a:pPr") {
		result = strings.ReplaceAll(result, "<a:p>", `<a:p><a:pPr algn="`+alVal+`"/>`)
		result = strings.ReplaceAll(result, "<a:p ", `<a:p><a:pPr algn="`+alVal+`"/> `)
		return result
	}

	for {
		idx := strings.Index(result, "<a:pPr")
		if idx < 0 {
			break
		}
		tagEnd := strings.Index(result[idx:], ">")
		if tagEnd < 0 {
			break
		}
		tagEnd += idx
		tagContent := result[idx+6 : tagEnd]

		// Check if it's self-closing
		selfClose := strings.HasSuffix(tagContent, "/")
		if selfClose {
			tagContent = tagContent[:len(tagContent)-1]
		}
		tagContent = strings.TrimSpace(tagContent)

		attrs := parseAttrString(tagContent)
		attrs["algn"] = alVal

		newTag := "<a:pPr"
		for k, v := range attrs {
			newTag += fmt.Sprintf(` %s="%s"`, k, v)
		}
		if selfClose {
			newTag += "/>"
		} else {
			newTag += ">"
		}

		result = result[:idx] + newTag + result[tagEnd+1:]
	}

	return result
}

// parseAttrString parses XML attributes from a tag content string like `lang="en-US" sz="1800"`.
func parseAttrString(s string) map[string]string {
	attrs := make(map[string]string)
	s = strings.TrimSpace(s)
	for s != "" {
		eqIdx := strings.Index(s, "=")
		if eqIdx < 0 {
			break
		}
		key := strings.TrimSpace(s[:eqIdx])
		s = s[eqIdx+1:]
		if len(s) == 0 {
			break
		}
		quote := s[0]
		if quote != '"' && quote != '\'' {
			break
		}
		s = s[1:]
		endIdx := strings.IndexByte(s, quote)
		if endIdx < 0 {
			break
		}
		attrs[key] = s[:endIdx]
		s = strings.TrimSpace(s[endIdx+1:])
	}
	return attrs
}

// ---------------------------------------------------------------------------
// AddShape — insert a new shape into a slide
// ---------------------------------------------------------------------------

// predefinedShapes maps ShapeSpec.Type to OOXML preset geometry names.
var predefinedShapes = map[string]string{
	"text-box": "rect",
	"rect":     "rect",
	"ellipse":  "ellipse",
	"oval":     "ellipse",
	"circle":   "ellipse",
	"line":     "line",
	"arrow":    "rightArrow",
}

// AddShape inserts a new shape into the specified slide.
func AddShape(path, outPath string, slideNum int, spec ShapeSpec) error {
	if slideNum < 1 {
		return fmt.Errorf("slide number must be >= 1, got %d", slideNum)
	}
	prst, ok := predefinedShapes[strings.ToLower(spec.Type)]
	if !ok {
		return fmt.Errorf("unsupported shape type %q (supported: text-box, rect, ellipse, line, arrow)", spec.Type)
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
	data, err := common.ReadEntry(zr, slideFile)
	if err != nil {
		_ = srcF.Close()
		return fmt.Errorf("slide %d not found", slideNum)
	}
	_ = srcF.Close()

	slideStr := string(data)

	// Find next available id by scanning existing <cNvPr id="N"> values
	nextID := maxCnvPrID(slideStr) + 1
	shapeName := fmt.Sprintf("Shape %d", nextID)

	// Build the shape XML
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf(`    <p:sp>
      <p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>
        <a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, nextID, shapeName, spec.X, spec.Y, spec.W, spec.H, prst))

	if spec.Fill != "" {
		sb.WriteString(fmt.Sprintf(`
        <a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, spec.Fill))
	}
	if spec.Line != "" {
		sb.WriteString(fmt.Sprintf(`
        <a:ln><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, spec.Line))
	}

	sb.WriteString(`
      </p:spPr>`)

	// Add text body unless it's a line (lines don't have text)
	if prst != "line" {
		fontSize := spec.FontSize
		if fontSize <= 0 {
			fontSize = 1800 // default 18pt
		}

		sb.WriteString(`
      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"`)
		sb.WriteString(fmt.Sprintf(` sz="%d"`, fontSize))
		if spec.Bold {
			sb.WriteString(` b="1"`)
		}
		sb.WriteString(">")
		if spec.Color != "" {
			sb.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, spec.Color))
		}
		sb.WriteString(fmt.Sprintf(`<a:t>%s</a:t></a:r></a:p></p:txBody>`, xmlEscapePPT(spec.Text)))
	}

	sb.WriteString(`
    </p:sp>`)

	shapeXML := sb.String()

	// Insert before </p:spTree> or </spTree>
	insertPoint := strings.LastIndex(slideStr, "</p:spTree>")
	if insertPoint < 0 {
		insertPoint = strings.LastIndex(slideStr, "</spTree>")
	}
	if insertPoint < 0 {
		return fmt.Errorf("invalid slide XML: missing </p:spTree>")
	}
	newSlide := slideStr[:insertPoint] + shapeXML + "\n" + slideStr[insertPoint:]

	out := outPath
	if out == "" {
		out = path
	}
	return common.RewriteEntries(path, out, map[string][]byte{
		slideFile: []byte(newSlide),
	})
}

// maxCnvPrID returns the highest id attribute value found in any <cNvPr> element.
func maxCnvPrID(xmlStr string) int {
	max := 0
	remaining := xmlStr
	for {
		idx := strings.Index(remaining, `<p:cNvPr `)
		if idx < 0 {
			break
		}
		// Also check without namespace prefix (for tests with default namespace)
		idx2 := strings.Index(remaining, `<cNvPr `)
		var nextIdx int
		if idx >= 0 && (idx2 < 0 || idx <= idx2) {
			nextIdx = idx
		} else if idx2 >= 0 {
			nextIdx = idx2
		} else {
			break
		}
		tagStart := nextIdx
		tagEnd := strings.Index(remaining[tagStart:], ">")
		if tagEnd < 0 {
			break
		}
		tagEnd += tagStart
		tag := remaining[tagStart:tagEnd]

		// Extract id="N"
		idIdx := strings.Index(tag, `id="`)
		if idIdx >= 0 {
			idStart := idIdx + 4
			idEnd := strings.Index(tag[idStart:], `"`)
			if idEnd >= 0 {
				if id, err := strconv.Atoi(tag[idStart : idStart+idEnd]); err == nil && id > max {
					max = id
				}
			}
		}
		remaining = remaining[tagEnd+1:]
	}
	return max
}
