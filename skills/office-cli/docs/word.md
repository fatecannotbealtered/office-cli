# Word (.docx) Command Reference

Permission levels: read/search/meta/stats/headings/images = `read-only`, replace/create/add-paragraph/add-heading/add-table/add-image/add-page-break/delete/insert-before/insert-after/update-table-cell/update-table/add-table-rows/delete-table-rows/add-table-cols/delete-table-cols/merge/style/style-table = `write`.

```bash
# Read paragraphs (default), heuristic markdown, or plain text
office-cli word read report.docx --format paragraphs --json
office-cli word read report.docx --format markdown --json
office-cli word read report.docx --format text --quiet

# Read paragraphs AND tables in document order (body elements)
office-cli word read report.docx --with-tables --json
office-cli word read report.docx --with-tables --format markdown --json
office-cli word read report.docx --with-tables --format text --json

# Filter to paragraphs containing a keyword
office-cli word read report.docx --keyword "revenue" --json

# Find/replace (creates a NEW file; the source is never mutated)
# NOTE: replace works on ALL text in document.xml — paragraphs AND table cells
office-cli word replace report.docx --find "TODO" --replace "DONE"
office-cli word replace report.docx --pairs '[
  {"find":"{{name}}","replace":"Alice"},
  {"find":"{{date}}","replace":"2026-05-04"}
]' --output report.filled.docx

# Search across all body elements (paragraphs + table cells)
office-cli word search report.docx --keyword "revenue" --json
# Returns: matching element index, type, and for tables: specific row/col/cell

# Inspect metadata (author, page count, word count, ...)
office-cli word meta report.docx --json

# Document statistics (paragraphs / headings / words / characters / lines)
office-cli word stats report.docx --json

# Just the heading outline (Title + Heading1..Heading6)
office-cli word headings report.docx --json

# Extract every embedded image into a directory
office-cli word images report.docx --output-dir ./images

# Create a brand-new document
office-cli word create report.docx --title "Quarterly Report" --author "Alice"

# Append paragraphs and headings
office-cli word add-paragraph report.docx --text "Introduction text here..." --output report.docx
office-cli word add-paragraph report.docx --text "Styled paragraph" --style "Normal" --output report.docx
office-cli word add-heading   report.docx --text "Section Title" --level 1 --output report.docx
office-cli word add-heading   report.docx --text "Subsection" --level 2 --output report.docx

# Append a table (first row = header)
office-cli word add-table report.docx --rows '[
  ["Name","Score","Grade"],
  ["Alice",95,"A"],
  ["Bob",82,"B"]
]' --output report.docx

# Insert an inline image (width/height in points; 72pt = 1 inch)
office-cli word add-image report.docx --image chart.png --width 400 --height 300 --output report.docx

# Insert a page break
office-cli word add-page-break report.docx --output report.docx

# Delete a body element (paragraph or table) by 0-based index
# Use --with-tables to discover the index of each element
office-cli word delete report.docx --index 3 --output report.trimmed.docx

# Insert a paragraph before/after a specific body element
office-cli word insert-before report.docx --index 0 --type paragraph --text "New opening" --output report.docx
office-cli word insert-after  report.docx --index 2 --type heading --text "New Section" --level 1 --output report.docx
office-cli word insert-before report.docx --index 1 --type table --rows '[["A","B"],["1","2"]]' --output report.docx

# Insert a page break at a specific position
office-cli word insert-before report.docx --index 3 --type page-break --output report.docx

# Update a single table cell
# table-index = body-element index of the table (from read --with-tables)
# row/col are 0-based
office-cli word update-table-cell report.docx --table-index 2 --row 1 --col 0 --value "Updated" --output report.docx

# Batch-update multiple cells in a table
office-cli word update-table report.docx --table-index 2 --spec '[
  {"row":0,"col":0,"value":"New A"},
  {"row":1,"col":2,"value":"New C"}
]' --output report.docx

# Add rows to an existing table
office-cli word add-table-rows report.docx --table-index 2 --rows '[["Charlie",78,"C"],["Dave",91,"A"]]' --output report.docx
# Insert at a specific position (before row index 1)
office-cli word add-table-rows report.docx --table-index 2 --position 1 --rows '[["New",0,"X"]]' --output report.docx

# Delete rows from a table (start-row and end-row are inclusive, 0-based)
office-cli word delete-table-rows report.docx --table-index 2 --start-row 2 --end-row 3 --output report.docx

# Add a column to a table
office-cli word add-table-cols report.docx --table-index 2 --values '["Bonus","100","50"]' --output report.docx
# Insert at a specific position (before column index 1)
office-cli word add-table-cols report.docx --table-index 2 --position 1 --values '["X","Y","Z"]' --output report.docx

# Delete columns from a table (start-col and end-col are inclusive, 0-based)
office-cli word delete-table-cols report.docx --table-index 2 --start-col 1 --end-col 2 --output report.docx

# Merge multiple .docx files into one
office-cli word merge intro.docx body.docx conclusion.docx --output full-report.docx

# Style a paragraph (element index from read --with-tables)
# Run formatting: --bold --italic --underline --strikethrough --font-family --font-size --color
# Paragraph formatting: --align --space-before --space-after --line-spacing --indent-left --indent-right --first-line
office-cli word style report.docx --index 0 --bold --color FF0000 --font-size 16 --align center --output report.docx

# Style table cells (by range — use -1 for "all")
office-cli word style-table report.docx --table-index 1 --start-row 0 --end-row 0 --bg-color 003366 --bold --color FFFFFF --output report.docx

# Style table cells with batch spec (multiple regions in one pass)
office-cli word style-table report.docx --table-index 1 --spec '[
  {"startRow":0,"endRow":0,"bgColor":"003366","bold":true,"color":"FFFFFF"},
  {"startRow":1,"bgColor":"FF0000","color":"FFFFFF"}
]' --output report.docx
```

## Structural editing workflow (read → modify → write)

For operations beyond simple append, the recommended Agent workflow is:

```bash
# 1. Discover the document structure
office-cli word read report.docx --with-tables --json

# 2. The output lists every body element with its 0-based index:
#    {"elements": [
#      {"index":0, "type":"paragraph", "style":"Title", "text":"Report"},
#      {"index":1, "type":"table", "rows":[["Name","Score"],["Alice","95"]]},
#      {"index":2, "type":"paragraph", "text":"Conclusion"}
#    ]}

# 3. Delete, insert, or update by index
office-cli word delete report.docx --index 1 --output report-edited.docx
office-cli word update-table-cell report.docx --table-index 1 --row 1 --col 1 --value "100" --output report-edited.docx
```

## Important caveats for AI Agents

1. **Run-split limitation:** Word stores text in "runs"; a single visible word can be split across multiple `<w:t>` elements when bold/italic/spell-check boundaries cut it. office-cli's `word replace` only matches text inside a single run. If a planned replacement does not produce hits, the target string is probably split. Use a unique placeholder token (e.g. `{{name}}`, `__SLOT__`) that is unlikely to be split.

2. **Replace works on both paragraphs and tables:** `word replace` operates on the raw bytes of `word/document.xml`, so it naturally hits text inside `<w:tbl>` elements (table cells) as well as paragraphs. The same run-split limitation applies to table cell text.

3. **Index stability:** After `delete` or `insert-before/after`, the 0-based indices of remaining elements shift. Always re-read with `--with-tables` after a structural mutation before issuing the next edit.

4. **Table identification:** Tables do not have names; they are identified solely by their body-element index. Use `read --with-tables` to discover the correct `--table-index` value before calling `update-table-cell`, `update-table`, `add-table-rows`, `delete-table-rows`, `add-table-cols`, or `delete-table-cols`.

5. **Merge preserves first file's styles:** `word merge` uses the first file as the base for styles, headers, and footers. Content from subsequent files is appended.

6. **Styling is additive:** `word style` and `word style-table` inject or replace XML style blocks (`<w:rPr>`, `<w:pPr>`, `<w:tcPr>`) without removing existing formatting that is not being changed. The `--index` / `--table-index` values come from `read --with-tables`.

7. **Font size is in points:** `--font-size 14` means 14pt (internally converted to 28 half-points as required by OOXML). Spacing and indentation values are in twips (1/20 of a point; 240 twips ≈ single line spacing).
