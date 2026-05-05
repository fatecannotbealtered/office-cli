// Package config manages the office-cli home directory, user configuration,
// and the three-tier permission system (read-only / write / full).
//
// Configuration is stored at ~/.office-cli/config.json. Environment variables
// override file-based settings for CI/CD use.
package config

import (
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// Permission levels — default is the most restrictive.
const (
	PermReadOnly = "read-only"
	PermWrite    = "write"
	PermFull     = "full"
)

// Config is the on-disk schema at ~/.office-cli/config.json.
type Config struct {
	Email       string            `json:"email,omitempty"`
	Server      string            `json:"server,omitempty"`
	Timezone    string            `json:"timezone,omitempty"`
	Permissions PermissionsConfig `json:"permissions"`
}

// PermissionsConfig holds the permission mode.
type PermissionsConfig struct {
	Mode string `json:"mode"`
}

// Dir returns the office-cli home directory: $OFFICE_CLI_HOME, or ~/.office-cli.
// The directory is NOT auto-created; callers create on demand.
func Dir() string {
	if h := os.Getenv("OFFICE_CLI_HOME"); h != "" {
		return h
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return ".office-cli"
	}
	return filepath.Join(home, ".office-cli")
}

// AuditDir returns the per-month audit log directory inside the home directory.
func AuditDir() string {
	return filepath.Join(Dir(), "audit")
}

// CacheDir returns the cache directory inside the home directory (reserved for future use).
func CacheDir() string {
	return filepath.Join(Dir(), "cache")
}

// ConfigPath returns the path to the config file.
func ConfigPath() string {
	return filepath.Join(Dir(), "config.json")
}

// Load reads the config file. Returns a default config if the file does not exist.
// Environment variables override file-based settings.
func Load() Config {
	cfg := Config{
		Permissions: PermissionsConfig{Mode: PermReadOnly},
		Timezone:    "Asia/Shanghai",
	}

	data, err := os.ReadFile(ConfigPath())
	if err == nil {
		_ = json.Unmarshal(data, &cfg)
	}

	// Ensure permissions default
	if cfg.Permissions.Mode == "" {
		cfg.Permissions.Mode = PermReadOnly
	}

	// Environment variable overrides
	if env := os.Getenv("OFFICE_CLI_PERMISSIONS"); env != "" {
		mode := NormalizePermMode(env)
		if mode != "" {
			cfg.Permissions.Mode = mode
		}
	}
	if tz := os.Getenv("OFFICE_CLI_TIMEZONE"); tz != "" {
		cfg.Timezone = tz
	}

	return cfg
}

// Save writes the config to disk. Creates the directory if needed.
// File permissions: 0600, directory: 0700.
func Save(cfg Config) error {
	dir := Dir()
	if err := os.MkdirAll(dir, 0700); err != nil {
		return fmt.Errorf("creating config directory: %w", err)
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return fmt.Errorf("marshaling config: %w", err)
	}
	data = append(data, '\n')
	path := ConfigPath()
	if err := os.WriteFile(path, data, 0600); err != nil {
		return fmt.Errorf("writing config: %w", err)
	}
	return nil
}

// NormalizePermMode normalizes a permission string to one of the canonical values.
// Returns "" if the input is not a valid permission mode.
func NormalizePermMode(s string) string {
	switch strings.ToLower(strings.TrimSpace(s)) {
	case "read-only", "readonly", "read_only":
		return PermReadOnly
	case "write":
		return PermWrite
	case "full":
		return PermFull
	default:
		return ""
	}
}

// PermLevel returns a numeric level for comparison: read-only=0, write=1, full=2.
func PermLevel(mode string) int {
	switch mode {
	case PermReadOnly:
		return 0
	case PermWrite:
		return 1
	case PermFull:
		return 2
	default:
		return 0
	}
}

// HasPermission reports whether the configured mode satisfies the required level.
func HasPermission(configured, required string) bool {
	return PermLevel(configured) >= PermLevel(required)
}

// Setup creates a default config file if one does not already exist.
func Setup() error {
	path := ConfigPath()
	if _, err := os.Stat(path); err == nil {
		return nil // already exists
	}
	cfg := Config{
		Permissions: PermissionsConfig{Mode: PermReadOnly},
		Timezone:    "Asia/Shanghai",
	}
	return Save(cfg)
}

// Validate checks that the config file is well-formed and readable.
// Returns a list of issues (empty if everything is OK).
func Validate() ([]string, error) {
	path := ConfigPath()
	if _, err := os.Stat(path); errors.Is(err, os.ErrNotExist) {
		return []string{"config file not found at " + path + "; run 'office-cli setup' to create it"}, nil
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("reading config: %w", err)
	}
	var cfg Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		return []string{"config file is not valid JSON: " + err.Error()}, nil
	}
	var issues []string
	mode := cfg.Permissions.Mode
	if mode == "" {
		issues = append(issues, "permissions.mode is empty; defaulting to read-only")
	} else if NormalizePermMode(mode) == "" {
		issues = append(issues, fmt.Sprintf("permissions.mode %q is not valid; expected read-only, write, or full", mode))
	}
	return issues, nil
}
