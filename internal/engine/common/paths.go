package common

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// FormatExtensions enumerates the file extensions recognized per format.
// Comparisons are always done after lower-casing.
var FormatExtensions = map[string][]string{
	"docx": {".docx"},
	"xlsx": {".xlsx", ".xlsm"},
	"pptx": {".pptx"},
	"pdf":  {".pdf"},
	"csv":  {".csv"},
	"md":   {".md", ".markdown"},
	"txt":  {".txt"},
}

// DetectFormat returns the canonical short name for a path, based on its
// extension. Returns "" when the extension is not recognized. Comparison is
// case-insensitive so "Report.DOCX" still resolves to "docx".
func DetectFormat(path string) string {
	ext := strings.ToLower(filepath.Ext(path))
	for name, exts := range FormatExtensions {
		for _, e := range exts {
			if e == ext {
				return name
			}
		}
	}
	return ""
}

// EnsureExtension fails loudly when path's extension does not belong to one of
// the allowed canonical formats. It is used to give AI Agents fast feedback.
func EnsureExtension(path string, formats ...string) error {
	got := DetectFormat(path)
	for _, want := range formats {
		if got == want {
			return nil
		}
	}
	allowed := strings.Join(formats, ", ")
	return fmt.Errorf("unsupported file extension for %s: expected one of [%s]", filepath.Base(path), allowed)
}

// NormalizePath cleans the user-provided path:
//   - expands a leading "~/" to the user's home directory
//   - converts to absolute (resolving relative paths against the cwd)
//   - applies filepath.Clean for cross-platform separator normalization
//
// Inputs containing spaces or non-ASCII characters (e.g. CJK) survive intact.
func NormalizePath(path string) (string, error) {
	if path == "" {
		return "", errors.New("empty path")
	}
	if strings.HasPrefix(path, "~/") || path == "~" {
		home, err := os.UserHomeDir()
		if err != nil {
			return "", fmt.Errorf("resolving ~: %w", err)
		}
		if path == "~" {
			path = home
		} else {
			path = filepath.Join(home, path[2:])
		}
	}
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", fmt.Errorf("resolving absolute path: %w", err)
	}
	return filepath.Clean(abs), nil
}

// EnsureParentDir creates the parent directory for the destination path if it
// does not yet exist. Used by every "writes a new file" command so that
// `--output dist/out.pdf` works even when `dist/` is absent.
func EnsureParentDir(path string) error {
	dir := filepath.Dir(path)
	if dir == "" || dir == "." {
		return nil
	}
	return os.MkdirAll(dir, 0755)
}

// SuggestSiblingPath returns a path next to the input file with `suffix`
// appended before the extension. E.g. SuggestSiblingPath("a.docx", ".replaced")
// returns "a.replaced.docx". Used when the user did not supply --output.
func SuggestSiblingPath(input, suffix string) string {
	ext := filepath.Ext(input)
	base := strings.TrimSuffix(input, ext)
	return base + suffix + ext
}
