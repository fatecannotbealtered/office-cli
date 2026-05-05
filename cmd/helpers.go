package cmd

import (
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/fatecannotbealtered/office-cli/internal/engine/common"
	"github.com/fatecannotbealtered/office-cli/internal/output"
)

// resolveInput is the canonical "open this file" gate used by every subcommand.
//
// It normalises the path (expanding "~", resolving relatives, cleaning separators),
// verifies the file exists, optionally enforces the file extension, and returns
// the absolute path back to the caller. On error it has already printed a
// structured response and the caller MUST return ErrSilent unchanged.
func resolveInput(path string, formats ...string) (string, error) {
	abs, err := common.NormalizePath(path)
	if err != nil {
		return "", emitError(err.Error(), output.ErrValidation, path, ExitBadArgs)
	}
	info, err := os.Stat(abs)
	if err != nil {
		return "", emitFileError(abs, err)
	}
	if info.IsDir() {
		return "", emitError("expected a file but found a directory: "+abs, output.ErrValidation, abs, ExitBadArgs)
	}
	if len(formats) > 0 {
		if err := common.EnsureExtension(abs, formats...); err != nil {
			return "", emitError(err.Error(), output.ErrInvalidFormat, abs, ExitBadArgs)
		}
	}
	return abs, nil
}

// resolveOutput normalises an output path. The file does NOT need to exist
// (it is being written to). The parent directory is created on demand.
func resolveOutput(path string) (string, error) {
	abs, err := common.NormalizePath(path)
	if err != nil {
		return "", emitError(err.Error(), output.ErrValidation, path, ExitBadArgs)
	}
	if err := common.EnsureParentDir(abs); err != nil {
		return "", emitError("creating parent directory: "+err.Error(), output.ErrPermission, abs, ExitForbidden)
	}
	return abs, nil
}

// readSpecArg reads JSON-typed flag values. Two forms are supported:
//   - inline JSON string:    --spec '{"sheets":...}'
//   - file reference:        --spec @path/to/spec.json
//
// The file form is preferred for AI Agents because shell escaping of large JSON
// payloads is brittle.
func readSpecArg(arg string) ([]byte, error) {
	if strings.HasPrefix(arg, "@") {
		path := strings.TrimPrefix(arg, "@")
		abs, err := common.NormalizePath(path)
		if err != nil {
			return nil, fmt.Errorf("resolving spec file: %w", err)
		}
		data, err := os.ReadFile(abs)
		if err != nil {
			return nil, fmt.Errorf("reading spec file: %w", err)
		}
		return common.StripBOM(data), nil
	}
	if strings.TrimSpace(arg) == "" {
		return nil, errors.New("spec is empty")
	}
	return []byte(arg), nil
}

// columnLetter converts a 1-based column index into its A1-style letter (1->A, 27->AA).
func columnLetter(n int) string {
	out := ""
	for n > 0 {
		n--
		out = string(rune('A'+n%26)) + out
		n /= 26
	}
	return out
}
