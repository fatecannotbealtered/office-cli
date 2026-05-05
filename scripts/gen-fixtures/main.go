// Command gen-fixtures generates real .docx / .xlsx / .pptx / .pdf documents
// for the integration test suite. The output is deterministic and contains
// CJK characters so the tests also exercise UTF-8 paths.
//
// Usage: go run ./scripts/gen-fixtures --out ./tmp/fixtures
//
// The generator deliberately avoids any external office library beyond
// excelize (already a project dependency). docx and pptx are written as
// minimal OOXML zips; PDF is written as a tiny hand-crafted document.
package main

import (
	"archive/zip"
	"bytes"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func main() {
	out := flag.String("out", "tmp/fixtures", "Directory where fixtures are written")
	flag.Parse()

	if err := os.MkdirAll(*out, 0755); err != nil {
		log.Fatal(err)
	}

	must("xlsx", makeXLSX(filepath.Join(*out, "sample.xlsx")))
	must("xlsx-cjk", makeXLSX(filepath.Join(*out, "样本.xlsx")))
	must("csv", makeCSV(filepath.Join(*out, "sample.csv")))
	must("docx", makeDOCX(filepath.Join(*out, "sample.docx")))
	must("pptx", makePPTX(filepath.Join(*out, "sample.pptx")))
	must("pdf", makePDF(filepath.Join(*out, "sample.pdf")))
	must("pdf-2", makePDFTwo(filepath.Join(*out, "sample-2.pdf")))

	fmt.Printf("fixtures written to %s\n", *out)
}

func must(label string, err error) {
	if err != nil {
		log.Fatalf("[%s] %v", label, err)
	}
}

// makeXLSX writes a workbook with an English sheet and a Chinese sheet
// containing values, a formula and an obvious search target.
func makeXLSX(path string) error {
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	first := f.GetSheetName(0)
	if first != "Sales" {
		if _, err := f.NewSheet("Sales"); err != nil {
			return err
		}
		if err := f.DeleteSheet(first); err != nil {
			return err
		}
	}
	rows := [][]any{
		{"Name", "Sales"},
		{"Alice", 100},
		{"Bob", 200},
		{"Carol", 300},
	}
	for r, row := range rows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellValue("Sales", ref, val); err != nil {
				return err
			}
		}
	}
	if err := f.SetCellFormula("Sales", "B5", "=SUM(B2:B4)"); err != nil {
		return err
	}

	if _, err := f.NewSheet("中文表"); err != nil {
		return err
	}
	zhRows := [][]any{
		{"姓名", "金额"},
		{"张三", 1234.5},
		{"李四", 5678.9},
	}
	for r, row := range zhRows {
		for c, val := range row {
			ref, _ := excelize.CoordinatesToCellName(c+1, r+1)
			if err := f.SetCellValue("中文表", ref, val); err != nil {
				return err
			}
		}
	}
	return f.SaveAs(path)
}

// makeCSV writes a small UTF-8 CSV (no BOM) with CJK characters so that the
// from-csv import path is exercised on a non-ASCII file.
func makeCSV(path string) error {
	body := "姓名,金额\n张三,1234.5\n李四,5678.9\n"
	return os.WriteFile(path, []byte(body), 0644)
}

// makeDOCX writes a minimal valid OOXML Word document. The body contains:
//   - Title style paragraph
//   - Heading1 paragraph
//   - Two normal paragraphs
//   - One paragraph with a {{NAME}} placeholder for replace tests
func makeDOCX(path string) error {
	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,

		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,

		"word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,

		"word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
</w:styles>`,

		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">office-cli 集成测试样例</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">This document is generated by gen-fixtures for end-to-end testing.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>背景</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">本段包含中文字符以验证 UTF-8 编码处理是否正确。</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Hello {{NAME}}, your invoice number is {{INVOICE}}.</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Score</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Grade</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>95</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>Bob</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>82</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`,

		"docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>office-cli sample</dc:title>
  <dc:creator>gen-fixtures</dc:creator>
  <dc:subject>integration test</dc:subject>
  <dc:description>Auto-generated fixture for office-cli</dc:description>
</cp:coreProperties>`,

		"docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>office-cli gen-fixtures</Application>
  <Pages>1</Pages>
  <Words>42</Words>
  <Characters>200</Characters>
</Properties>`,
	}
	return writeZip(path, files)
}

// makePPTX writes a 3-slide minimal pptx. The first slide has a title placeholder
// (so title detection works), other slides have body text.
func makePPTX(path string) error {
	slide := func(title, body string) string {
		titlePart := ""
		if title != "" {
			titlePart = `
        <p:sp>
          <p:nvSpPr><p:cNvPr id="1" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
          <p:spPr/>
          <p:txBody><a:bodyPr/><a:p><a:r><a:t>` + title + `</a:t></a:r></a:p></p:txBody>
        </p:sp>`
		}
		return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>` + titlePart + `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>` + body + `</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
	}

	files := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,

		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,

		"ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
</Relationships>`,

		"ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
    <p:sldId id="258" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
</p:presentation>`,

		"ppt/slides/slide1.xml": slide("office-cli 演示", "Hello world from gen-fixtures"),
		"ppt/slides/slide2.xml": slide("Roadmap", "Q1 foundation, Q2 GA, {{MILESTONE}}"),
		"ppt/slides/slide3.xml": slide("中文标题", "本幻灯片包含中文，用于验证 UTF-8 处理"),

		"docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>office-cli pptx sample</dc:title>
  <dc:creator>gen-fixtures</dc:creator>
</cp:coreProperties>`,

		"docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>office-cli gen-fixtures</Application>
  <Slides>3</Slides>
</Properties>`,
	}
	return writeZip(path, files)
}

// writeZip is the shared zip writer used by docx/pptx generators.
func writeZip(path string, files map[string]string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	w := zip.NewWriter(f)
	for name, content := range files {
		hdr := &zip.FileHeader{Name: name, Method: zip.Deflate}
		entry, err := w.CreateHeader(hdr)
		if err != nil {
			return err
		}
		if _, err := entry.Write([]byte(content)); err != nil {
			return err
		}
	}
	return w.Close()
}

// makePDF writes a tiny single-page PDF that says "Hello office-cli".
// We hand-roll it because pulling in a PDF authoring lib for fixtures alone
// would bloat the project. The PDF format is verbose but well-defined.
func makePDF(path string) error {
	return writeMinimalPDF(path, []string{"Hello office-cli"})
}

// makePDFTwo writes a 2-page PDF used by merge / split / trim tests.
func makePDFTwo(path string) error {
	return writeMinimalPDF(path, []string{"Page A: hello", "Page B: world"})
}

// writeMinimalPDF emits a deterministic PDF with one page per `pages` element.
// Each page contains the page text rendered with Helvetica 24pt at (100, 700).
//
// Implementation notes:
//   - We assemble the body bytes first, recording each object's byte offset.
//   - We then emit the xref table using those offsets and the trailer.
//   - The result passes "Validated" in pdfcpu and is rendered correctly by the
//     pdf reader and most PDF viewers.
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
	objects := []obj{}

	addObj := func(num int, content string) {
		objects = append(objects, obj{num: num, offset: body.Len()})
		fmt.Fprintf(&body, "%d 0 obj\n%s\nendobj\n", num, content)
	}

	pageObjStart := 3
	contentObjStart := pageObjStart + len(pages)
	fontObjNum := contentObjStart + len(pages)

	pagesKids := ""
	for i := range pages {
		if i > 0 {
			pagesKids += " "
		}
		pagesKids += fmt.Sprintf("%d 0 R", pageObjStart+i)
	}

	addObj(1, "<< /Type /Catalog /Pages 2 0 R >>")
	addObj(2, fmt.Sprintf("<< /Type /Pages /Kids [%s] /Count %d >>", pagesKids, len(pages)))

	for i, text := range pages {
		pageNum := pageObjStart + i
		contentNum := contentObjStart + i
		pageDict := fmt.Sprintf("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents %d 0 R /Resources << /Font << /F1 %d 0 R >> >> >>", contentNum, fontObjNum)
		addObj(pageNum, pageDict)
		stream := fmt.Sprintf("BT /F1 24 Tf 100 700 Td (%s) Tj ET\n", escapePDFString(text))
		streamObj := fmt.Sprintf("<< /Length %d >>\nstream\n%sendstream", len(stream), stream)
		addObj(contentNum, streamObj)
	}
	addObj(fontObjNum, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

	xrefOffset := body.Len()
	totalObjects := fontObjNum + 1
	body.WriteString("xref\n")
	fmt.Fprintf(&body, "0 %d\n", totalObjects)
	body.WriteString("0000000000 65535 f \n")
	for n := 1; n < totalObjects; n++ {
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
	fmt.Fprintf(&body, "trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n",
		totalObjects, xrefOffset)

	return os.WriteFile(path, body.Bytes(), 0644)
}

// escapePDFString escapes the printable subset we use. We avoid binary content
// so the result is readable and stable across runs.
func escapePDFString(s string) string {
	var b bytes.Buffer
	for _, r := range s {
		switch r {
		case '(', ')', '\\':
			b.WriteByte('\\')
			b.WriteRune(r)
		default:
			if r > 0x7f {
				// PDF Type1 fonts can't render Unicode without ToUnicode CMap;
				// we simply replace non-ASCII with '?' for the fixture body.
				b.WriteByte('?')
				continue
			}
			b.WriteRune(r)
		}
	}
	return b.String()
}
