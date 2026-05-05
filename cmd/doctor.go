package cmd

import (
	"fmt"
	"os/exec"
	"runtime"

	"github.com/fatecannotbealtered/office-cli/internal/config"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var doctorCmd = &cobra.Command{
	Use:   "doctor",
	Short: "Check runtime environment and optional external tools",
	Long: `Inspect the local environment for office-cli. Reports Go runtime info,
office-cli version, and the availability of optional external tools used for
advanced format conversions:

  pandoc        — Word ↔ Markdown / HTML / LaTeX conversions
  libreoffice   — Office documents → PDF (when 'office-cli word convert ... --to pdf' is requested)

These tools are NOT required for the core read / write / search / merge commands.
The CLI works fully offline using built-in Go libraries.`,
	RunE: runDoctor,
}

func init() {
	rootCmd.AddCommand(doctorCmd)
}

// toolStatus is the JSON shape for one external-tool entry.
type toolStatus struct {
	Name      string `json:"name"`
	Found     bool   `json:"found"`
	Path      string `json:"path,omitempty"`
	UsedFor   string `json:"usedFor"`
	InstallOS string `json:"installHint,omitempty"`
}

// optionalTools lists every tool that office-cli can leverage when present.
// All operations remain functional without them; we just downgrade gracefully.
var optionalTools = []toolStatus{
	{Name: "pandoc", UsedFor: "Word/Markdown/HTML conversions"},
	{Name: "libreoffice", UsedFor: "Headless export to PDF (Word, PPT)"},
	{Name: "soffice", UsedFor: "Alternative LibreOffice binary name on some platforms"},
}

func runDoctor(_ *cobra.Command, _ []string) error {
	type configResult struct {
		Path   string   `json:"path"`
		Mode   string   `json:"mode"`
		Status string   `json:"status"`
		Issues []string `json:"issues,omitempty"`
	}

	type result struct {
		Version string       `json:"version"`
		GoVer   string       `json:"goVersion"`
		OS      string       `json:"os"`
		Arch    string       `json:"arch"`
		Tools   []toolStatus `json:"tools"`
		Config  configResult `json:"config"`
	}

	cfg := config.Load()
	cfgIssues, _ := config.Validate()
	cfgStatus := "ok"
	if len(cfgIssues) > 0 {
		cfgStatus = "issues"
	}

	r := result{
		Version: version,
		GoVer:   runtime.Version(),
		OS:      runtime.GOOS,
		Arch:    runtime.GOARCH,
		Tools:   make([]toolStatus, 0, len(optionalTools)),
		Config: configResult{
			Path:   config.ConfigPath(),
			Mode:   cfg.Permissions.Mode,
			Status: cfgStatus,
			Issues: cfgIssues,
		},
	}

	for _, t := range optionalTools {
		path, err := exec.LookPath(t.Name)
		t.Found = err == nil
		t.Path = path
		t.InstallOS = installHint(t.Name)
		r.Tools = append(r.Tools, t)
	}

	if jsonMode {
		output.PrintJSON(r)
		return nil
	}

	fmt.Println()
	output.Bold("  office-cli Doctor")
	output.Gray("  ────────────────────────────────────────")
	fmt.Println()
	output.Success(fmt.Sprintf("office-cli %s (Go %s, %s/%s)", r.Version, r.GoVer, r.OS, r.Arch))
	fmt.Println()

	// Config status
	output.Bold("  Configuration:")
	if r.Config.Status == "ok" {
		output.Success(fmt.Sprintf("config valid: %s (mode: %s)", r.Config.Path, r.Config.Mode))
	} else {
		output.Warn(fmt.Sprintf("config issues at %s (mode: %s)", r.Config.Path, r.Config.Mode))
		for _, issue := range r.Config.Issues {
			output.Warn("    - " + issue)
		}
	}
	fmt.Println()

	output.Bold("  Optional tools:")
	rows := make([][]string, 0, len(r.Tools))
	for _, t := range r.Tools {
		mark := output.FormatRed("missing")
		path := "—"
		if t.Found {
			mark = output.FormatGreen("found")
			path = t.Path
		}
		rows = append(rows, []string{t.Name, mark, t.UsedFor, path})
	}
	output.Table([]string{"tool", "status", "purpose", "path"}, rows)
	fmt.Println()
	output.Gray("  Built-in operations work fully offline; missing tools only disable optional flags.")
	fmt.Println()
	return nil
}

// installHint returns a one-line install suggestion per tool. Best-effort, OS-aware.
func installHint(tool string) string {
	switch tool {
	case "pandoc":
		switch runtime.GOOS {
		case "darwin":
			return "brew install pandoc"
		case "linux":
			return "apt install pandoc  (or yum / dnf / pacman equivalent)"
		case "windows":
			return "choco install pandoc  (or winget install pandoc)"
		}
	case "libreoffice", "soffice":
		switch runtime.GOOS {
		case "darwin":
			return "brew install --cask libreoffice"
		case "linux":
			return "apt install libreoffice"
		case "windows":
			return "Download from https://www.libreoffice.org/"
		}
	}
	return ""
}
