# PowerPoint (.pptx) Command Reference

Permission levels: read/count/outline/meta/images/layout = `read-only`, replace/create/add-slide/set-content/set-notes/delete-slide/reorder/add-image/build/set-style/add-shape = `write`.

```bash
# Read the outline (titles + bullets)
office-cli ppt read deck.pptx --json
office-cli ppt read deck.pptx --slide 3 --json
office-cli ppt read deck.pptx --keyword "roadmap" --with-notes --json
office-cli ppt read deck.pptx --format markdown               # render as Markdown outline
office-cli ppt read deck.pptx --format text                   # everything as plain text

# Find/replace across every slide
office-cli ppt replace deck.pptx --find "{{client}}" --replace "Acme Inc."
office-cli ppt replace deck.pptx --pairs @placeholders.json --output deck.final.pptx

# Inspect metadata (author, slide count, application, ...)
office-cli ppt meta deck.pptx --json

# Lightweight calls when you just need "how many slides" or the title outline
office-cli ppt count   deck.pptx --json
office-cli ppt outline deck.pptx --json

# Extract every embedded image into a directory
office-cli ppt images deck.pptx --output-dir ./slide-images

# Create a brand-new presentation
office-cli ppt create deck.pptx --title "Q2 Review" --author "Alice"

# Append a new slide with title and bullets
office-cli ppt add-slide deck.pptx --title "Roadmap" --bullets '["Q1: foundation","Q2: GA","Q3: scale"]' --output deck.pptx

# Overwrite a slide's text content
office-cli ppt set-content deck.pptx --slide 3 --title "Updated Roadmap" --bullets '["Revised timeline","New milestones"]' --output deck.pptx

# Set / replace speaker notes for a slide
office-cli ppt set-notes deck.pptx --slide 3 --notes "Skip if running short on time" --output deck.pptx

# Delete a slide by number
office-cli ppt delete-slide deck.pptx --slide 2 --output deck.pptx

# Reorder slides (comma-separated new order)
office-cli ppt reorder deck.pptx --order "3,1,2,4" --output deck.pptx

# Insert an image into a specific slide
office-cli ppt add-image deck.pptx --slide 1 --image logo.png --output deck.pptx

# Create a complete deck from a JSON spec in one shot
office-cli ppt build deck.pptx --spec '{"title":"Q2 Review","author":"Alice","slides":[{"title":"Overview","bullets":["Revenue up 20%"]},{"title":"Q&A","notes":"Thank the team"}]}'
office-cli ppt build deck.pptx --spec @deck.json --json

# Build a deck with images embedded in specific slides
office-cli ppt build deck.pptx --spec '{"slides":[{"title":"Chart","image":"chart.png","width":6000000,"height":4000000},{"title":"Summary"}]}'

# Build a deck from a template (preserves slide master, theme, layouts)
office-cli ppt build deck.pptx --template company.pptx --spec '{"title":"Q2 Report","slides":[{"title":"Overview"},{"title":"Details"},{"title":"Q&A"}]}'

# --- Layout & styling ---

# Read the shape tree of a slide (position, size, type, text, font info)
office-cli ppt layout deck.pptx --slide 1 --json
office-cli ppt layout deck.pptx --json          # all slides

# Modify text styling within a shape (use shape index from 'ppt layout')
office-cli ppt set-style deck.pptx --slide 1 --shape 0 --font-size 4800 --bold --color "FF0000" --output deck.pptx
office-cli ppt set-style deck.pptx --slide 1 --shape 0 --align center --output deck.pptx

# Insert a new shape into a slide
office-cli ppt add-shape deck.pptx --slide 1 --type text-box --x 500000 --y 200000 --width 4000000 --height 1000000 --text "Hello World" --font-size 2400 --fill "E8E8E8" --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type rect --x 0 --y 0 --width 2000000 --height 1000000 --fill "4472C4" --line "000000" --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type ellipse --x 1000000 --y 1000000 --width 3000000 --height 2000000 --text "Oval" --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type line --x 0 --y 0 --width 9144000 --height 0 --output deck.pptx
office-cli ppt add-shape deck.pptx --slide 1 --type arrow --x 500000 --y 500000 --width 3000000 --height 0 --output deck.pptx
```

The same multi-run caveat from Word applies. Prefer placeholder tokens (`{{client}}`, `__SLOT__`) for reliable replacement.
