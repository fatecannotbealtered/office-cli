package audit

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestSanitizeArgs(t *testing.T) {
	cases := []struct {
		name string
		in   []string
		want []string
	}{
		{
			name: "no sensitive flags",
			in:   []string{"excel", "read", "file.xlsx", "--json"},
			want: []string{"excel", "read", "file.xlsx", "--json"},
		},
		{
			name: "strips --token",
			in:   []string{"setup", "--token", "secret123", "--host", "jira.example.com"},
			want: []string{"setup", "--host", "jira.example.com"},
		},
		{
			name: "strips --password",
			in:   []string{"--password", "hunter2", "--force"},
			want: []string{"--force"},
		},
		{
			name: "strips --secret",
			in:   []string{"--secret", "abc", "arg1"},
			want: []string{"arg1"},
		},
		{
			name: "strips --user-password",
			in:   []string{"pdf", "encrypt", "file.pdf", "--user-password", "1234", "--output", "out.pdf"},
			want: []string{"pdf", "encrypt", "file.pdf", "--output", "out.pdf"},
		},
		{
			name: "strips --owner-password",
			in:   []string{"--owner-password", "admin", "--user-password", "user", "--force"},
			want: []string{"--force"},
		},
		{
			name: "case insensitive",
			in:   []string{"--Token", "val", "--PASSWORD", "val2"},
			want: []string{},
		},
		{
			name: "empty input",
			in:   []string{},
			want: []string{},
		},
		{
			name: "sensitive flag at end (no value to skip)",
			in:   []string{"--token"},
			want: []string{},
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			got := sanitizeArgs(tc.in)
			if len(got) != len(tc.want) {
				t.Fatalf("len mismatch: got %d, want %d (%v vs %v)", len(got), len(tc.want), got, tc.want)
			}
			for i := range got {
				if got[i] != tc.want[i] {
					t.Errorf("index %d: got %q, want %q", i, got[i], tc.want[i])
				}
			}
		})
	}
}

func TestLogWritesEntry(t *testing.T) {
	dir := t.TempDir()
	testDir = dir
	defer func() { testDir = "" }()

	Log("office-cli test cmd", []string{"arg1", "--flag", "val"}, 0, 42)

	files, err := Files()
	if err != nil {
		t.Fatalf("Files(): %v", err)
	}
	if len(files) != 1 {
		t.Fatalf("expected 1 audit file, got %d", len(files))
	}

	data, err := os.ReadFile(files[0])
	if err != nil {
		t.Fatalf("read audit file: %v", err)
	}

	var e entry
	if err := json.Unmarshal(data, &e); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if e.Cmd != "office-cli test cmd" {
		t.Errorf("cmd = %q, want %q", e.Cmd, "office-cli test cmd")
	}
	if e.Exit != 0 {
		t.Errorf("exit = %d, want 0", e.Exit)
	}
	if e.Ms != 42 {
		t.Errorf("ms = %d, want 42", e.Ms)
	}
	if e.TS == "" {
		t.Error("ts should not be empty")
	}
}

func TestLogDisabledByEnv(t *testing.T) {
	dir := t.TempDir()
	testDir = dir
	defer func() { testDir = "" }()

	t.Setenv("OFFICE_CLI_NO_AUDIT", "1")
	Log("cmd", []string{"arg"}, 0, 10)

	files, err := Files()
	if err != nil {
		t.Fatalf("Files(): %v", err)
	}
	if len(files) != 0 {
		t.Errorf("expected no audit files when disabled, got %d", len(files))
	}
}

func TestCleanup(t *testing.T) {
	dir := t.TempDir()

	// Create a fake old audit file
	oldFile := filepath.Join(dir, "audit-2020-01.jsonl")
	if err := os.WriteFile(oldFile, []byte("{}\n"), 0600); err != nil {
		t.Fatal(err)
	}
	// Create a recent file
	recentFile := filepath.Join(dir, "audit-2026-05.jsonl")
	if err := os.WriteFile(recentFile, []byte("{}\n"), 0600); err != nil {
		t.Fatal(err)
	}

	cleanup(dir)

	if _, err := os.Stat(oldFile); !os.IsNotExist(err) {
		t.Error("old audit file should have been deleted")
	}
	if _, err := os.Stat(recentFile); err != nil {
		t.Error("recent audit file should NOT have been deleted")
	}
}

func TestRetentionMonths(t *testing.T) {
	// Default
	if got := retentionMonths(); got != 3 {
		t.Errorf("default retention = %d, want 3", got)
	}

	t.Setenv("OFFICE_CLI_AUDIT_RETENTION_MONTHS", "6")
	if got := retentionMonths(); got != 6 {
		t.Errorf("retention = %d, want 6", got)
	}

	t.Setenv("OFFICE_CLI_AUDIT_RETENTION_MONTHS", "0")
	if got := retentionMonths(); got != 0 {
		t.Errorf("retention = %d, want 0", got)
	}

	t.Setenv("OFFICE_CLI_AUDIT_RETENTION_MONTHS", "bad")
	if got := retentionMonths(); got != 3 {
		t.Errorf("bad value: retention = %d, want 3", got)
	}
}

func TestFilesEmptyDir(t *testing.T) {
	dir := t.TempDir()
	testDir = dir
	defer func() { testDir = "" }()

	files, err := Files()
	if err != nil {
		t.Fatalf("Files(): %v", err)
	}
	if len(files) != 0 {
		t.Errorf("expected 0 files, got %d", len(files))
	}
}
