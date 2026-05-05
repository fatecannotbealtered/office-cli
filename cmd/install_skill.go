package cmd

import (
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

var installSkillCmd = &cobra.Command{
	Use:   "install-skill",
	Short: "Install the bundled AI Agent Skill into ~/.agent/skills/",
	Long: `Copies the bundled SKILL.md (and any companion reference files) from the binary
distribution into ~/.agent/skills/ so that AI Agent platforms (Cursor, Claude
Code, OpenClaw, etc.) can discover and use office-cli automatically.`,
	RunE: runInstallSkill,
}

func init() {
	rootCmd.AddCommand(installSkillCmd)
}

// findSkillsDir resolves the bundled skills directory. Order:
//  1. <binary-dir>/skills           (release tarball layout)
//  2. <binary-dir>/../skills        (npm: packageRoot/bin/<bin> + packageRoot/skills)
//  3. ./skills                      (local dev with `go run` from repo root)
func findSkillsDir(execDir string) string {
	candidates := []string{
		filepath.Join(execDir, "skills"),
		filepath.Join(execDir, "..", "skills"),
	}
	for _, dir := range candidates {
		if fi, err := os.Stat(dir); err == nil && fi.IsDir() {
			return dir
		}
	}
	if fi, err := os.Stat("skills"); err == nil && fi.IsDir() {
		return "skills"
	}
	return ""
}

func runInstallSkill(cmd *cobra.Command, _ []string) error {
	execPath, err := os.Executable()
	if err != nil {
		return emitError("failed to locate binary: "+err.Error(), output.ErrEngine, "", ExitEngine)
	}
	execDir := filepath.Dir(execPath)
	skillsDir := findSkillsDir(execDir)
	if skillsDir == "" {
		return emitError("no skill files found to install", output.ErrValidation, "", ExitBadArgs)
	}

	home, err := os.UserHomeDir()
	if err != nil {
		return emitError("failed to get home directory: "+err.Error(), output.ErrEngine, "", ExitEngine)
	}
	targetDir := filepath.Join(home, ".agent", "skills")
	if err := os.MkdirAll(targetDir, 0755); err != nil {
		return emitError("failed to create target directory: "+err.Error(), output.ErrPermission, targetDir, ExitForbidden)
	}

	var installedFiles []string
	err = filepath.Walk(skillsDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info.IsDir() {
			return nil
		}
		rel, _ := filepath.Rel(skillsDir, path)
		destPath := filepath.Join(targetDir, rel)

		if err := os.MkdirAll(filepath.Dir(destPath), 0755); err != nil {
			return err
		}

		_, existErr := os.Stat(destPath)
		updated := existErr == nil

		if err := copyFile(path, destPath); err != nil {
			return fmt.Errorf("copying %s: %w", rel, err)
		}

		installedFiles = append(installedFiles, rel)
		if updated {
			output.Info(fmt.Sprintf("Updated: %s", rel))
		} else {
			output.Success(fmt.Sprintf("Installed: %s", rel))
		}
		return nil
	})

	if err != nil {
		return emitError("install failed: "+err.Error(), output.ErrEngine, "", ExitEngine)
	}

	if len(installedFiles) == 0 {
		return emitError("no skill files found to install", output.ErrValidation, "", ExitBadArgs)
	}

	if jsonMode {
		output.PrintJSON(map[string]any{
			"installedFiles": installedFiles,
			"targetDir":      targetDir,
		})
		return nil
	}

	fmt.Println()
	output.Success(fmt.Sprintf("Skill installed to %s", targetDir))
	output.Gray("  AI Agents will now have access to office-cli capabilities.")
	fmt.Println()
	return nil
}

func copyFile(src, dst string) error {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer func() { _ = in.Close() }()

	out, err := os.Create(dst)
	if err != nil {
		return err
	}
	defer func() { _ = out.Close() }()

	_, err = io.Copy(out, in)
	return err
}
