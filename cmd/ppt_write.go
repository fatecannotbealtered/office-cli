package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/fatecannotbealtered/office-cli/internal/engine/ppt"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// ---------------------------------------------------------------------------
// New PPT write commands: create, add-slide, set-content, set-notes, delete-slide, reorder, add-image
// ---------------------------------------------------------------------------

var pptCreateCmd = &cobra.Command{
	Use:   "create <FILE>",
	Short: "Create a new minimal .pptx presentation",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path := args[0]
		title, _ := cmd.Flags().GetString("title")
		author, _ := cmd.Flags().GetString("author")

		if dryRunOutput("create presentation", map[string]any{"file": path, "title": title, "author": author}) {
			return nil
		}

		if err := ppt.Create(path, title, author); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path})
			return nil
		}
		output.Success(fmt.Sprintf("created %s", path))
		return nil
	},
}

var pptAddSlideCmd = &cobra.Command{
	Use:   "add-slide <FILE>",
	Short: "Append a new slide with optional title and bullets",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		title, _ := cmd.Flags().GetString("title")
		bulletsRaw, _ := cmd.Flags().GetString("bullets")
		out, _ := cmd.Flags().GetString("output")

		var bullets []string
		if bulletsRaw != "" {
			data, err := readSpecArg(bulletsRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &bullets); err != nil {
				return emitError("invalid --bullets JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add slide", map[string]any{"file": path, "title": title}) {
			return nil
		}

		if err := ppt.AddSlide(path, out, title, bullets); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out})
			return nil
		}
		output.Success(fmt.Sprintf("added slide to %s", out))
		return nil
	},
}

var pptSetContentCmd = &cobra.Command{
	Use:   "set-content <FILE>",
	Short: "Overwrite a slide's title and bullets",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		title, _ := cmd.Flags().GetString("title")
		bulletsRaw, _ := cmd.Flags().GetString("bullets")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs, "Use 'ppt count' to check total slides, then specify --slide 1..N")
		}

		var bullets []string
		if bulletsRaw != "" {
			data, err := readSpecArg(bulletsRaw)
			if err != nil {
				return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
			if err := json.Unmarshal(data, &bullets); err != nil {
				return emitError("invalid --bullets JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs)
			}
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("set slide content", map[string]any{"file": path, "slide": slideNum}) {
			return nil
		}

		if err := ppt.SetSlideContent(path, out, slideNum, title, bullets); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "slide": slideNum})
			return nil
		}
		output.Success(fmt.Sprintf("updated slide %d in %s", slideNum, out))
		return nil
	},
}

var pptSetNotesCmd = &cobra.Command{
	Use:   "set-notes <FILE>",
	Short: "Set speaker notes for a slide",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		notes, _ := cmd.Flags().GetString("notes")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if notes == "" {
			return emitError("--notes is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("set notes", map[string]any{"file": path, "slide": slideNum}) {
			return nil
		}

		if err := ppt.SetNotes(path, out, slideNum, notes); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "slide": slideNum})
			return nil
		}
		output.Success(fmt.Sprintf("set notes for slide %d in %s", slideNum, out))
		return nil
	},
}

var pptDeleteSlideCmd = &cobra.Command{
	Use:   "delete-slide <FILE>",
	Short: "Remove a slide by number",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("delete slide", map[string]any{"file": path, "slide": slideNum}) {
			return nil
		}

		if err := ppt.DeleteSlide(path, out, slideNum); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "deleted": slideNum})
			return nil
		}
		output.Success(fmt.Sprintf("deleted slide %d -> %s", slideNum, out))
		return nil
	},
}

var pptReorderCmd = &cobra.Command{
	Use:   "reorder <FILE>",
	Short: "Reorder slides by providing a comma-separated order (e.g. '3,1,2')",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		order, _ := cmd.Flags().GetString("order")
		out, _ := cmd.Flags().GetString("output")

		if order == "" {
			return emitError("--order is required (e.g. '3,1,2')", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("reorder slides", map[string]any{"file": path, "order": order}) {
			return nil
		}

		if err := ppt.ReorderSlides(path, out, order); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "order": order})
			return nil
		}
		output.Success(fmt.Sprintf("reordered slides -> %s", out))
		return nil
	},
}

var pptAddImageCmd = &cobra.Command{
	Use:   "add-image <FILE>",
	Short: "Insert an image into a specific slide",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		imagePath, _ := cmd.Flags().GetString("image")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if imagePath == "" {
			return emitError("--image is required", output.ErrValidation, "", ExitBadArgs)
		}
		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add image", map[string]any{"file": path, "slide": slideNum, "image": imagePath}) {
			return nil
		}

		widthEMU, _ := cmd.Flags().GetInt("width")
		heightEMU, _ := cmd.Flags().GetInt("height")
		if err := ppt.AddImage(path, out, slideNum, imagePath, widthEMU, heightEMU); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "slide": slideNum, "image": imagePath})
			return nil
		}
		output.Success(fmt.Sprintf("added image to slide %d in %s", slideNum, out))
		return nil
	},
}

var pptBuildCmd = &cobra.Command{
	Use:   "build <FILE>",
	Short: "Create a complete deck from a JSON spec in one shot",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		specRaw, _ := cmd.Flags().GetString("spec")
		if specRaw == "" {
			return emitError("--spec is required (JSON string or @file.json)", output.ErrValidation, "", ExitBadArgs, "Provide a JSON string or @file.json path with title, author, and slides array")
		}
		data, err := readSpecArg(specRaw)
		if err != nil {
			return emitError(err.Error(), output.ErrValidation, "", ExitBadArgs, "Ensure --spec is valid JSON or a path to a .json file prefixed with @")
		}
		var spec ppt.BuildSpec
		if err := json.Unmarshal(data, &spec); err != nil {
			return emitError("invalid --spec JSON: "+err.Error(), output.ErrValidation, "", ExitBadArgs, "Check JSON syntax: slides must be an array of objects with title, bullets, notes, image fields")
		}
		if len(spec.Slides) == 0 {
			return emitError("slides array must not be empty", output.ErrValidation, "", ExitBadArgs, "Provide at least one slide object in the slides array")
		}
		template, _ := cmd.Flags().GetString("template")
		path := args[0]
		if dryRunOutput("build deck", map[string]any{"file": path, "slides": len(spec.Slides), "title": spec.Title, "template": template}) {
			return nil
		}
		if template != "" {
			if err := ppt.BuildFromTemplate(template, path, spec); err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine, "Verify the template file exists and is a valid .pptx with at least one slide")
			}
		} else {
			if err := ppt.Build(path, spec); err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine, "Ensure all image paths in --spec are valid and accessible")
			}
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": path, "slides": len(spec.Slides)})
			return nil
		}
		output.Success(fmt.Sprintf("built %s (%d slides)", path, len(spec.Slides)))
		return nil
	},
}

// ---------------------------------------------------------------------------
// Phase 3: layout, set-style, add-shape
// ---------------------------------------------------------------------------

var pptLayoutCmd = &cobra.Command{
	Use:   "layout <FILE>",
	Short: "Read the shape tree of a slide (position, size, type, text, font info)",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")

		if slideNum > 0 {
			shapes, err := ppt.ReadSlideLayout(path, slideNum)
			if err != nil {
				return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
			}
			if jsonMode {
				output.PrintJSON(map[string]any{"file": path, "slide": slideNum, "shapes": shapes})
				return nil
			}
			// Human-friendly table
			rows := make([][]string, 0, len(shapes))
			for _, s := range shapes {
				ph := s.Ph
				if ph == "" {
					ph = "-"
				}
				rows = append(rows, []string{
					fmt.Sprintf("%d", s.Index),
					s.Type,
					s.Name,
					ph,
					fmt.Sprintf("(%d,%d) %dx%d", s.X, s.Y, s.W, s.H),
					s.Text,
				})
			}
			output.Table([]string{"#", "type", "name", "ph", "pos/size", "text"}, rows)
			return nil
		}

		// All slides
		count, err := ppt.SlideCount(path)
		if err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		allShapes := map[int][]ppt.ShapeInfo{}
		for i := 1; i <= count; i++ {
			shapes, err := ppt.ReadSlideLayout(path, i)
			if err != nil {
				continue
			}
			allShapes[i] = shapes
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"file": path, "slides": allShapes})
			return nil
		}
		for i := 1; i <= count; i++ {
			shapes := allShapes[i]
			fmt.Printf("--- Slide %d (%d shapes) ---\n", i, len(shapes))
			rows := make([][]string, 0, len(shapes))
			for _, s := range shapes {
				ph := s.Ph
				if ph == "" {
					ph = "-"
				}
				rows = append(rows, []string{
					fmt.Sprintf("%d", s.Index),
					s.Type,
					s.Name,
					ph,
					fmt.Sprintf("(%d,%d) %dx%d", s.X, s.Y, s.W, s.H),
					s.Text,
				})
			}
			output.Table([]string{"#", "type", "name", "ph", "pos/size", "text"}, rows)
		}
		return nil
	},
}

var pptSetStyleCmd = &cobra.Command{
	Use:   "set-style <FILE>",
	Short: "Modify text styling (font size, bold, color, alignment) within a shape",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		shapeIdx, _ := cmd.Flags().GetInt("shape")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if shapeIdx < 0 {
			return emitError("--shape is required (0-based, from 'ppt layout')", output.ErrValidation, "", ExitBadArgs)
		}

		fontSize, _ := cmd.Flags().GetInt("font-size")
		bold, _ := cmd.Flags().GetBool("bold")
		italic, _ := cmd.Flags().GetBool("italic")
		underline, _ := cmd.Flags().GetBool("underline")
		color, _ := cmd.Flags().GetString("color")
		align, _ := cmd.Flags().GetString("align")

		opts := ppt.StyleOptions{}
		changed := false
		if cmd.Flags().Changed("font-size") {
			opts.FontSize = &fontSize
			changed = true
		}
		if cmd.Flags().Changed("bold") {
			opts.Bold = &bold
			changed = true
		}
		if cmd.Flags().Changed("italic") {
			opts.Italic = &italic
			changed = true
		}
		if cmd.Flags().Changed("underline") {
			opts.Underline = &underline
			changed = true
		}
		if color != "" {
			opts.Color = &color
			changed = true
		}
		if align != "" {
			opts.Align = &align
			changed = true
		}
		if !changed {
			return emitError("provide at least one style flag (--font-size, --bold, --color, --align, etc.)", output.ErrValidation, "", ExitBadArgs)
		}

		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("set-style", map[string]any{"file": path, "slide": slideNum, "shape": shapeIdx}) {
			return nil
		}

		if err := ppt.SetShapeStyle(path, out, slideNum, shapeIdx, opts); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "slide": slideNum, "shape": shapeIdx})
			return nil
		}
		output.Success(fmt.Sprintf("styled shape %d on slide %d -> %s", shapeIdx, slideNum, out))
		return nil
	},
}

var pptAddShapeCmd = &cobra.Command{
	Use:   "add-shape <FILE>",
	Short: "Insert a new shape (text-box, rect, ellipse, line, arrow) into a slide",
	Args:  cobra.ExactArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		path, err := resolveInput(args[0], "pptx")
		if err != nil {
			return err
		}
		slideNum, _ := cmd.Flags().GetInt("slide")
		shapeType, _ := cmd.Flags().GetString("type")
		out, _ := cmd.Flags().GetString("output")

		if slideNum < 1 {
			return emitError("--slide is required (1-based)", output.ErrValidation, "", ExitBadArgs)
		}
		if shapeType == "" {
			return emitError("--type is required (text-box, rect, ellipse, line, arrow)", output.ErrValidation, "", ExitBadArgs)
		}

		x, _ := cmd.Flags().GetInt("x")
		y, _ := cmd.Flags().GetInt("y")
		w, _ := cmd.Flags().GetInt("width")
		h, _ := cmd.Flags().GetInt("height")
		text, _ := cmd.Flags().GetString("text")
		fontSize, _ := cmd.Flags().GetInt("font-size")
		bold, _ := cmd.Flags().GetBool("bold")
		color, _ := cmd.Flags().GetString("color")
		fill, _ := cmd.Flags().GetString("fill")
		line, _ := cmd.Flags().GetString("line")

		spec := ppt.ShapeSpec{
			Type:     shapeType,
			X:        x,
			Y:        y,
			W:        w,
			H:        h,
			Text:     text,
			FontSize: fontSize,
			Bold:     bold,
			Color:    color,
			Fill:     fill,
			Line:     line,
		}

		if out == "" {
			out = path
		} else {
			out, err = resolveOutput(out)
			if err != nil {
				return err
			}
		}

		if dryRunOutput("add-shape", map[string]any{"file": path, "slide": slideNum, "type": shapeType}) {
			return nil
		}

		if err := ppt.AddShape(path, out, slideNum, spec); err != nil {
			return emitError(err.Error(), output.ErrEngine, path, ExitEngine)
		}
		if jsonMode {
			output.PrintJSON(map[string]any{"status": "ok", "file": out, "slide": slideNum, "type": shapeType})
			return nil
		}
		output.Success(fmt.Sprintf("added %s to slide %d -> %s", shapeType, slideNum, out))
		return nil
	},
}
