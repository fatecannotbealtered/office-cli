# PDF Command Reference

Permission levels: read/pages/info/bookmarks/search = `read-only`, merge/split/trim/watermark/rotate/optimize/extract-images/reorder/insert-blank/stamp-image/set-meta/add-bookmarks/create/replace/add-text = `write`, encrypt/decrypt = `full`.

```bash
# Page count and (optionally) per-page dimensions
office-cli pdf pages report.pdf --json
office-cli pdf pages report.pdf --dimensions --json

# Read text — per page or as a single concatenated string
office-cli pdf read report.pdf --json                        # all pages, structured
office-cli pdf read report.pdf --page 5 --json               # only page 5
office-cli pdf read report.pdf --from 2 --to 4 --json
office-cli pdf read report.pdf --text-only --quiet           # everything as one string

# Search for text in a PDF (case-insensitive by default)
office-cli pdf search report.pdf "keyword" --json
office-cli pdf search report.pdf "keyword" --page 3 --json   # restrict to page 3
office-cli pdf search report.pdf "keyword" --case-sensitive --context 10 --limit 20 --json

# Create a new PDF from a JSON spec
office-cli pdf create output.pdf --spec '{"pages":[{"text":"Hello World"},{"text":"Page 2"}]}' --json
office-cli pdf create output.pdf --spec '{"pages":[{"text":"Content"}],"title":"My Doc","author":"Alice","paper":"A4"}' --json

# Create with typography options
office-cli pdf create output.pdf --spec '{"pages":[{"text":"Serif text","font":"Times","bold":true}],"font":"Times","fontSize":14}' --json
office-cli pdf create output.pdf --spec '{"pages":[{"text":"Centered","align":"center"},{"text":"Right aligned","align":"right"}]}' --json
office-cli pdf create output.pdf --spec '{"pages":[{"text":"Underlined","underline":true},{"text":"Monospace","font":"Courier"}]}' --json
office-cli pdf create output.pdf --spec '{"font":"Times","align":"center","lineHeight":1.8,"margin":50,"pages":[{"text":"Global defaults"}]}' --json

# Find and replace text in a PDF
office-cli pdf replace report.pdf --find "Old Text" --replace "New Text" --output report.fixed.pdf
office-cli pdf replace report.pdf --pairs '[{"find":"Alice","replace":"Bob"},{"find":"X","replace":"Y"}]' --output report.fixed.pdf

# Merge multiple PDFs (positional args or --input flags both work)
office-cli pdf merge a.pdf b.pdf c.pdf --output combined.pdf
office-cli pdf merge --input a.pdf --input b.pdf --output combined.pdf

# Split into chunks of N pages each
office-cli pdf split big.pdf --span 10 --output-dir ./parts/

# Keep only specific pages (Trim)
office-cli pdf trim big.pdf --pages "1,3,5-7,10" --output excerpt.pdf

# Watermark every page
office-cli pdf watermark report.pdf --text "DRAFT" --output report.draft.pdf
office-cli pdf watermark report.pdf --text "Confidential" --style "font:Helvetica, points:36, opacity:0.4, rotation:30, color:0.8 0.0 0.0" --output report.confidential.pdf

# Full metadata (encryption, signatures, dimensions, watermark / form / tagged flags)
office-cli pdf info report.pdf --json

# Rotate selected pages (90 / 180 / 270 degrees)
office-cli pdf rotate report.pdf --degrees 90 --pages "1,3-4" --output rotated.pdf

# Optimize file size by compacting cross-references and reusing objects
office-cli pdf optimize report.pdf --output smaller.pdf

# Encrypt / decrypt
office-cli pdf encrypt report.pdf --user-password 1234 --output secret.pdf
office-cli pdf decrypt secret.pdf --user-password 1234 --output report.plain.pdf

# Extract every embedded image into a directory
office-cli pdf extract-images report.pdf --output-dir ./pdf-images

# Read PDF bookmarks / outline
office-cli pdf bookmarks report.pdf --json

# Add bookmarks to a PDF
office-cli pdf add-bookmarks report.pdf --spec '[
  {"title":"Introduction","page":1},
  {"title":"Chapter 1","page":3,"children":[{"title":"Section 1.1","page":4}]}
]' --output report.bookmarked.pdf

# Reorder pages (comma-separated page numbers in desired order)
office-cli pdf reorder report.pdf --order "3,1,2,4" --output reordered.pdf

# Insert blank pages at a position
office-cli pdf insert-blank report.pdf --after 2 --count 1 --output report.with-blanks.pdf

# Stamp an image watermark on pages
office-cli pdf stamp-image report.pdf --image logo.png --pages "1-" --output report.stamped.pdf

# Set PDF metadata
office-cli pdf set-meta report.pdf --title "Annual Report" --author "Alice" --subject "Finance" --keywords "2026,annual" --output report.meta.pdf

# Add text overlay at specific coordinates
office-cli pdf add-text report.pdf --text "DRAFT" --x 200 --y 400 --font-size 48 --color "1 0 0" --output report.overlay.pdf
office-cli pdf add-text report.pdf --text "Page 1 stamp" --x 72 --y 700 --pages "1" --output report.overlay.pdf

# Add text with font options
office-cli pdf add-text report.pdf --text "Serif Bold" --font Times --bold --italic --x 200 --y 400 --output report.overlay.pdf
office-cli pdf add-text report.pdf --text "Underlined" --underline --x 72 --y 700 --output report.overlay.pdf

# Add multiple text overlays via --spec
office-cli pdf add-text report.pdf --spec '[
  {"text":"CONFIDENTIAL","x":200,"y":400,"fontSize":36,"color":"1 0 0"},
  {"text":"Draft v1.0","x":72,"y":50,"fontSize":10,"pages":"1"}
]' --output report.overlay.pdf

# Add multiple text overlays with font options via --spec
office-cli pdf add-text report.pdf --spec '[
  {"text":"Serif Title","x":200,"y":700,"fontSize":24,"font":"Times","bold":true},
  {"text":"Underlined Note","x":72,"y":50,"fontSize":10,"underline":true,"pages":"1"}
]' --output report.overlay.pdf
```

PDF text extraction works on content-stream text (typed PDFs). Scanned/image-only PDFs need an OCR pipeline first; office-cli does not OCR.

### pdf create spec schema

```json
{
  "pages": [
    {"text": "Page 1 content. Newlines are\nrespected and long lines wrap automatically."},
    {"text": "Bold red heading", "fontSize": 18, "bold": true, "color": "1 0 0"},
    {"text": "Small blue note", "fontSize": 8, "color": "0 0 1"},
    {"text": "Serif text", "font": "Times", "italic": true},
    {"text": "Centered monospace", "font": "Courier", "align": "center"},
    {"text": "Underlined", "underline": true}
  ],
  "title": "Document Title",
  "author": "Author Name",
  "paper": "A4",
  "font": "Times",
  "fontSize": 14,
  "align": "center",
  "lineHeight": 1.6,
  "margin": 72
}
```

**Page-level fields:**
- `text` (required): page content. Newlines are respected and long lines wrap automatically.
- `fontSize`: font size in points (default 12)
- `bold`: use bold font variant (default false)
- `italic`: use italic font variant (default false)
- `underline`: add underline decoration (default false)
- `color`: text color as `"R G B"` in 0–1 range (e.g. `"1 0 0"` = red, `"0 0 1"` = blue)
- `font`: font family — `"Helvetica"` (default, sans-serif), `"Times"` (serif), or `"Courier"` (monospace)
- `align`: text alignment — `"left"` (default), `"center"`, or `"right"`
- `lineHeight`: line height multiplier (default 1.4)
- `margin`: page margin in points (default 72, i.e. 1 inch)

**Global fields (applied to all pages unless overridden):**
- `font`: default font family for all pages
- `fontSize`: default font size for all pages
- `align`: default alignment for all pages
- `lineHeight`: default line height for all pages
- `margin`: default margin for all pages

**Other fields:**
- `paper`: `"Letter"` (default, 612×792pt), `"A4"` (595×842pt), or `"WxH"` (custom, e.g. `"400x600"`)
- `title`: PDF metadata title
- `author`: PDF metadata author

### pdf replace limitations

`pdf replace` modifies text stored as parenthesized string literals `(text)` and hex-encoded strings `<hex>` in the PDF content stream. It does NOT handle:
- Text split across multiple Tj operators
- Text rendered via font-specific encodings that don't map to ASCII

For these cases, use `word replace` on the source `.docx` before converting to PDF.

### pdf add-text spec schema

```json
[
  {"text": "DRAFT", "x": 200, "y": 400, "fontSize": 36, "color": "1 0 0"},
  {"text": "Page 1 only", "x": 72, "y": 700, "pages": "1"},
  {"text": "Serif Bold", "x": 200, "y": 300, "font": "Times", "bold": true, "italic": true},
  {"text": "Underlined", "x": 72, "y": 50, "underline": true, "pages": "1"}
]
```

- `text` (required): the text to overlay
- `x`, `y` (required): position in points from bottom-left corner
- `fontSize`: font size in points (default 12)
- `bold`: use bold font variant (default false)
- `italic`: use italic font variant (default false)
- `underline`: add underline decoration (default false)
- `color`: `"R G B"` in 0–1 range (e.g. `"1 0 0"` = red)
- `font`: font family — `"Helvetica"` (default, sans-serif), `"Times"` (serif), or `"Courier"` (monospace)
- `pages`: page selection (e.g. `"1,3,5-7"`); empty = all pages
