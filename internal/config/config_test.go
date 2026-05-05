package config

import (
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestNormalizePermMode(t *testing.T) {
	cases := []struct {
		in   string
		want string
	}{
		{"read-only", PermReadOnly},
		{"readonly", PermReadOnly},
		{"read_only", PermReadOnly},
		{"Read-Only", PermReadOnly},
		{"write", PermWrite},
		{"Write", PermWrite},
		{"full", PermFull},
		{"FULL", PermFull},
		{"invalid", ""},
		{"", ""},
		{" admin ", ""},
	}
	for _, tc := range cases {
		t.Run(tc.in, func(t *testing.T) {
			got := NormalizePermMode(tc.in)
			if got != tc.want {
				t.Errorf("NormalizePermMode(%q) = %q, want %q", tc.in, got, tc.want)
			}
		})
	}
}

func TestPermLevel(t *testing.T) {
	if PermLevel(PermReadOnly) != 0 {
		t.Error("read-only should be 0")
	}
	if PermLevel(PermWrite) != 1 {
		t.Error("write should be 1")
	}
	if PermLevel(PermFull) != 2 {
		t.Error("full should be 2")
	}
	if PermLevel("unknown") != 0 {
		t.Error("unknown should default to 0")
	}
}

func TestHasPermission(t *testing.T) {
	cases := []struct {
		configured, required string
		want                 bool
	}{
		{PermReadOnly, PermReadOnly, true},
		{PermReadOnly, PermWrite, false},
		{PermReadOnly, PermFull, false},
		{PermWrite, PermReadOnly, true},
		{PermWrite, PermWrite, true},
		{PermWrite, PermFull, false},
		{PermFull, PermReadOnly, true},
		{PermFull, PermWrite, true},
		{PermFull, PermFull, true},
	}
	for _, tc := range cases {
		t.Run(tc.configured+"/"+tc.required, func(t *testing.T) {
			got := HasPermission(tc.configured, tc.required)
			if got != tc.want {
				t.Errorf("HasPermission(%q, %q) = %v, want %v", tc.configured, tc.required, got, tc.want)
			}
		})
	}
}

func TestLoadDefault(t *testing.T) {
	// Point to a non-existent directory to get defaults
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", filepath.Join(dir, "nonexistent"))

	cfg := Load()
	if cfg.Permissions.Mode != PermReadOnly {
		t.Errorf("default mode = %q, want %q", cfg.Permissions.Mode, PermReadOnly)
	}
	if cfg.Timezone != "Asia/Shanghai" {
		t.Errorf("default timezone = %q, want %q", cfg.Timezone, "Asia/Shanghai")
	}
}

func TestLoadFromFile(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	cfgPath := filepath.Join(dir, "config.json")
	content := `{"email":"test@example.com","timezone":"UTC","permissions":{"mode":"write"}}`
	if err := os.WriteFile(cfgPath, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}

	cfg := Load()
	if cfg.Email != "test@example.com" {
		t.Errorf("email = %q, want %q", cfg.Email, "test@example.com")
	}
	if cfg.Permissions.Mode != PermWrite {
		t.Errorf("mode = %q, want %q", cfg.Permissions.Mode, PermWrite)
	}
	if cfg.Timezone != "UTC" {
		t.Errorf("timezone = %q, want %q", cfg.Timezone, "UTC")
	}
}

func TestLoadEnvOverride(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	cfgPath := filepath.Join(dir, "config.json")
	content := `{"permissions":{"mode":"read-only"}}`
	if err := os.WriteFile(cfgPath, []byte(content), 0600); err != nil {
		t.Fatal(err)
	}

	t.Setenv("OFFICE_CLI_PERMISSIONS", "full")
	cfg := Load()
	if cfg.Permissions.Mode != PermFull {
		t.Errorf("mode = %q, want %q (env should override file)", cfg.Permissions.Mode, PermFull)
	}
}

func TestSaveAndLoad(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	cfg := Config{
		Email:       "user@test.com",
		Timezone:    "America/New_York",
		Permissions: PermissionsConfig{Mode: PermFull},
	}
	if err := Save(cfg); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// Verify file permissions (skip on Windows — ACLs, not Unix perms)
	if runtime.GOOS != "windows" {
		info, err := os.Stat(ConfigPath())
		if err != nil {
			t.Fatalf("Stat: %v", err)
		}
		if info.Mode().Perm() != 0600 {
			t.Errorf("file mode = %o, want 0600", info.Mode().Perm())
		}
	}

	loaded := Load()
	if loaded.Email != "user@test.com" {
		t.Errorf("email = %q, want %q", loaded.Email, "user@test.com")
	}
	if loaded.Permissions.Mode != PermFull {
		t.Errorf("mode = %q, want %q", loaded.Permissions.Mode, PermFull)
	}
}

func TestSetup(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	if err := Setup(); err != nil {
		t.Fatalf("Setup: %v", err)
	}

	if _, err := os.Stat(ConfigPath()); err != nil {
		t.Fatalf("config file not created: %v", err)
	}

	// Running again should be a no-op
	if err := Setup(); err != nil {
		t.Fatalf("Setup (second call): %v", err)
	}
}

func TestValidateNoFile(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	issues, err := Validate()
	if err != nil {
		t.Fatalf("Validate: %v", err)
	}
	if len(issues) == 0 {
		t.Error("expected issues when config file doesn't exist")
	}
}

func TestValidateValidFile(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	cfg := Config{Permissions: PermissionsConfig{Mode: PermWrite}}
	if err := Save(cfg); err != nil {
		t.Fatal(err)
	}

	issues, err := Validate()
	if err != nil {
		t.Fatalf("Validate: %v", err)
	}
	if len(issues) != 0 {
		t.Errorf("expected no issues, got %v", issues)
	}
}

func TestValidateInvalidJSON(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	if err := os.WriteFile(ConfigPath(), []byte("not json"), 0600); err != nil {
		t.Fatal(err)
	}

	issues, err := Validate()
	if err != nil {
		t.Fatalf("Validate: %v", err)
	}
	if len(issues) == 0 {
		t.Error("expected issues for invalid JSON")
	}
}

func TestValidateInvalidMode(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", dir)

	cfg := Config{Permissions: PermissionsConfig{Mode: "admin"}}
	if err := Save(cfg); err != nil {
		t.Fatal(err)
	}

	issues, err := Validate()
	if err != nil {
		t.Fatalf("Validate: %v", err)
	}
	if len(issues) == 0 {
		t.Error("expected issues for invalid mode")
	}
}

func TestEnvOverridePermission(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("OFFICE_CLI_HOME", filepath.Join(dir, "nonexistent"))
	t.Setenv("OFFICE_CLI_PERMISSIONS", "bad-value")

	cfg := Load()
	// Bad env value should not override default
	if cfg.Permissions.Mode != PermReadOnly {
		t.Errorf("bad env: mode = %q, want %q", cfg.Permissions.Mode, PermReadOnly)
	}
}
