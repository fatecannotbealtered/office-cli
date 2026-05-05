// Package cmd defines the cobra command tree for office-cli.
//
// The root command exposes four global flags shared by every subcommand:
//
//	--json     machine-readable output (errors on stderr, results on stdout)
//	--quiet    suppress non-JSON stdout (pipe-friendly for AI Agents)
//	--dry-run  show what a write command would do without touching the filesystem
//	--force    skip interactive confirmations (CI / Agent automation)
//
// Subcommands organized by document family (word, excel, ppt, pdf) live in their
// own files but share a common error / exit-code / audit pipeline declared here.
package cmd

import (
	"errors"
	"fmt"
	"os"
	"runtime/debug"
	"time"

	"github.com/fatecannotbealtered/office-cli/internal/audit"
	"github.com/fatecannotbealtered/office-cli/internal/config"
	"github.com/fatecannotbealtered/office-cli/internal/output"
	"github.com/spf13/cobra"
)

// Stable, human-meaningful exit codes used by main().
const (
	ExitOK        = 0
	ExitBadArgs   = 2
	ExitNotFound  = 4
	ExitForbidden = 5
	ExitEngine    = 7
)

// Permission level annotation keys.
const (
	permAnnotation = "permission"
	permRead       = "read-only"
	permWrite      = "write"
	permFull       = "full"
)

// ErrSilent signals that the error has already been rendered (cobra should not re-print it).
var ErrSilent = errors.New("")

// version is overridden via -ldflags by Makefile / goreleaser.
var version = "dev"

// Global flag values, populated by cobra in init().
var (
	jsonMode  bool
	forceMode bool
	quietMode bool
	dryRun    bool
)

// lastExit tracks the worst exit code seen during the current invocation.
var lastExit int

// cmdStartTime is captured by PersistentPreRunE for audit duration.
var cmdStartTime time.Time

// loadedConfig holds the runtime configuration including permissions.
var loadedConfig config.Config

// LastExitCode is consumed by main() after Execute() returns ErrSilent.
func LastExitCode() int { return lastExit }

// setExitCode raises the recorded exit code (severity is monotonically non-decreasing).
func setExitCode(code int) {
	if code > lastExit {
		lastExit = code
	}
}

var rootCmd = &cobra.Command{
	Use:           "office-cli",
	Short:         "Local Office documents CLI for humans and AI Agents",
	Version:       version,
	SilenceErrors: true,
	SilenceUsage:  true,
	Long: fmt.Sprintf("\n  %s\n  %s",
		output.FormatCyanBold("office-cli"),
		output.FormatGray("Read, write, search and convert Word / Excel / PPT / PDF files locally")),
}

func init() {
	if version == "dev" {
		if info, ok := debug.ReadBuildInfo(); ok && info.Main.Version != "" {
			version = info.Main.Version
		}
	}
	rootCmd.Version = version
	rootCmd.CompletionOptions.DisableDefaultCmd = true

	rootCmd.PersistentFlags().BoolVar(&jsonMode, "json", false, "Output result as JSON (machine-readable)")
	rootCmd.PersistentFlags().BoolVar(&forceMode, "force", false, "Skip confirmation prompts")
	rootCmd.PersistentFlags().BoolVar(&quietMode, "quiet", false, "Suppress non-JSON stdout (for scripts and AI Agents)")
	rootCmd.PersistentFlags().BoolVar(&dryRun, "dry-run", false, "Show what would be done without writing the filesystem")

	cobra.OnInitialize(func() {
		output.Quiet = quietMode
	})

	rootCmd.PersistentPreRunE = func(cmd *cobra.Command, args []string) error {
		cmdStartTime = time.Now()
		loadedConfig = config.Load()

		// Permission check: skip for commands that don't need it (help, version, etc.)
		if cmd.CommandPath() == rootCmd.CommandPath() {
			return nil
		}
		required := cmdPermissionLevel(cmd)
		if !config.HasPermission(loadedConfig.Permissions.Mode, required) {
			msg := fmt.Sprintf("permission denied: command %q requires %q, but configured mode is %q",
				cmd.CommandPath(), required, loadedConfig.Permissions.Mode)
			if jsonMode {
				output.PrintErrorJSON(msg, output.ErrPermission)
			} else {
				output.Error(msg)
			}
			setExitCode(ExitForbidden)
			return ErrSilent
		}
		return nil
	}

	rootCmd.PersistentPostRunE = func(cmd *cobra.Command, args []string) error {
		if !isWriteCommand(cmd) {
			return nil
		}
		duration := time.Since(cmdStartTime)
		audit.Log(cmd.CommandPath(), os.Args[1:], lastExit, duration.Milliseconds())
		return nil
	}
}

// Execute runs the root command. main.go translates the result into a process exit code.
func Execute() error {
	return rootCmd.Execute()
}

// emitError renders err with the right channel/format and sets the exit code.
// In --json mode the structured error goes to stderr; otherwise a colored line.
// Returns ErrSilent so cobra leaves the message alone.
// An optional hint override can be provided; if non-empty, it replaces the
// default hint for the given error code (useful for context-specific guidance).
func emitError(msg string, code output.ErrorCode, file string, exit int, hint ...string) error {
	if jsonMode {
		output.PrintErrorJSONWithFile(msg, code, file, hint...)
	} else {
		if file != "" {
			output.Error(fmt.Sprintf("%s (%s)", msg, file))
		} else {
			output.Error(msg)
		}
	}
	setExitCode(exit)
	return ErrSilent
}

// emitFileError is a helper for the (very common) "open file failed" case.
// It maps os errors to user-visible messages and appropriate exit codes.
func emitFileError(file string, err error) error {
	switch {
	case os.IsNotExist(err):
		return emitError("file not found: "+file, output.ErrFileNotFound, file, ExitNotFound)
	case os.IsPermission(err):
		return emitError("permission denied: "+file, output.ErrPermission, file, ExitForbidden)
	default:
		return emitError(err.Error(), output.ErrEngine, file, ExitEngine)
	}
}

// dryRunOutput emits an informational line / JSON when --dry-run is set, returning true.
// Returns false in normal mode so the caller proceeds with the actual write.
func dryRunOutput(action string, detail map[string]any) bool {
	if !dryRun {
		return false
	}
	if jsonMode {
		if detail == nil {
			detail = map[string]any{}
		}
		detail["action"] = action
		detail["dryRun"] = true
		output.PrintJSON(detail)
	} else {
		output.Info("[dry-run] " + action)
	}
	return true
}

// confirmAction asks the user to type expected to confirm a destructive operation.
// Returns true immediately when --force is set, allowing CI / Agents to bypass.
func confirmAction(prompt, expected string) bool {
	if forceMode {
		return true
	}
	fmt.Printf("%s: ", prompt)
	var input string
	_, err := fmt.Fscan(os.Stdin, &input)
	if err != nil {
		return false
	}
	return input == expected
}

// isWriteCommand reports whether cmd has been tagged with markWrite() — used to gate audit logging.
func isWriteCommand(cmd *cobra.Command) bool {
	lvl := cmd.Annotations[permAnnotation]
	return lvl == permWrite || lvl == permFull
}

// markWrite tags cmd as a write-class command (permission level: write).
func markWrite(cmd *cobra.Command) {
	markPerm(cmd, permWrite)
}

// markFull tags cmd as requiring full permissions (irreversible operations).
func markFull(cmd *cobra.Command) {
	markPerm(cmd, permFull)
}

// markPerm sets the permission annotation on cmd.
func markPerm(cmd *cobra.Command, level string) {
	if cmd.Annotations == nil {
		cmd.Annotations = map[string]string{}
	}
	cmd.Annotations[permAnnotation] = level
}

// cmdPermissionLevel returns the required permission level for cmd.
// Defaults to read-only if not annotated.
func cmdPermissionLevel(cmd *cobra.Command) string {
	if lvl, ok := cmd.Annotations[permAnnotation]; ok {
		return lvl
	}
	return permRead
}
