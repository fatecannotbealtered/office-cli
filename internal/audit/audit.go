// Package audit records every write-class command (create, update, write, append,
// merge, split, watermark, ...) into a JSONL file under ~/.office-cli/audit/.
//
// The log is intended for traceability when AI Agents drive the CLI: each entry
// captures timestamp, command path, sanitized arguments, exit code and duration.
// Read-only commands (read, list, search, meta) are NOT audited.
package audit

import (
	"encoding/json"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/fatecannotbealtered/office-cli/internal/config"
)

// entry is the wire shape of one audit JSONL line.
type entry struct {
	TS   string   `json:"ts"`
	Cmd  string   `json:"cmd"`
	Args []string `json:"args"`
	Exit int      `json:"exit"`
	Ms   int64    `json:"ms"`
}

// testDir overrides Dir() during unit tests; not exported intentionally.
var testDir string

// Dir returns the audit log directory.
func Dir() string {
	if testDir != "" {
		return testDir
	}
	return config.AuditDir()
}

// Log appends one audit line for the just-completed command.
// No-op when OFFICE_CLI_NO_AUDIT=1 or when the directory cannot be created.
// Lazy-cleans audit files older than retentionMonths().
func Log(cmdPath string, args []string, exitCode int, durationMs int64) {
	if os.Getenv("OFFICE_CLI_NO_AUDIT") == "1" {
		return
	}

	dir := Dir()
	if err := os.MkdirAll(dir, 0700); err != nil {
		return
	}

	cleanup(dir)

	e := entry{
		TS:   time.Now().Format(time.RFC3339Nano),
		Cmd:  cmdPath,
		Args: sanitizeArgs(args),
		Exit: exitCode,
		Ms:   durationMs,
	}

	data, err := json.Marshal(e)
	if err != nil {
		return
	}
	data = append(data, '\n')

	filename := "audit-" + time.Now().Format("2006-01") + ".jsonl"
	path := filepath.Join(dir, filename)

	f, err := os.OpenFile(path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0600)
	if err != nil {
		return
	}
	defer func() { _ = f.Close() }()
	_, _ = f.Write(data)
}

// retentionMonths reads OFFICE_CLI_AUDIT_RETENTION_MONTHS (default 3).
// Returning 0 disables cleanup entirely.
func retentionMonths() int {
	s := os.Getenv("OFFICE_CLI_AUDIT_RETENTION_MONTHS")
	if s == "" {
		return 3
	}
	n, err := strconv.Atoi(s)
	if err != nil || n < 0 {
		return 3
	}
	return n
}

// cleanup deletes audit files whose YYYY-MM stamp is older than the cutoff.
func cleanup(dir string) {
	months := retentionMonths()
	if months == 0 {
		return
	}
	cutoff := time.Now().AddDate(0, -months, 0).Format("2006-01")

	entries, err := os.ReadDir(dir)
	if err != nil {
		return
	}
	for _, e := range entries {
		name := e.Name()
		if !strings.HasPrefix(name, "audit-") || !strings.HasSuffix(name, ".jsonl") {
			continue
		}
		ym := strings.TrimPrefix(name, "audit-")
		ym = strings.TrimSuffix(ym, ".jsonl")
		if ym < cutoff {
			_ = os.Remove(filepath.Join(dir, name))
		}
	}
}

// sanitizeArgs strips obviously sensitive flag values. office-cli has no auth
// today, but we redact any future --token / --password style flags defensively.
func sanitizeArgs(args []string) []string {
	out := make([]string, 0, len(args))
	skip := false
	for _, a := range args {
		if skip {
			skip = false
			continue
		}
		lower := strings.ToLower(a)
		if lower == "--token" || lower == "--password" || lower == "--secret" ||
			lower == "--user-password" || lower == "--owner-password" {
			skip = true
			continue
		}
		out = append(out, a)
	}
	return out
}

// Files returns the sorted list of audit JSONL files. Used by tests and tooling.
func Files() ([]string, error) {
	dir := Dir()
	entries, err := os.ReadDir(dir)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	var files []string
	for _, e := range entries {
		if !e.IsDir() && strings.HasSuffix(e.Name(), ".jsonl") {
			files = append(files, filepath.Join(dir, e.Name()))
		}
	}
	sort.Strings(files)
	return files, nil
}
