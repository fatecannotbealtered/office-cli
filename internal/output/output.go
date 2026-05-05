// Package output centralizes terminal output (colored human text, JSON, tables) for office-cli.
//
// The package is shared by all commands so that human and AI Agent consumers
// see consistent formatting, color, error structure and silencing behavior.
package output

import (
	"fmt"
	"os"
)

// ANSI color codes used for human-friendly stdout/stderr output.
const (
	ansiReset    = "\033[0m"
	ansiBold     = "\033[1m"
	ansiRed      = "\033[31m"
	ansiGreen    = "\033[32m"
	ansiYellow   = "\033[33m"
	ansiBlue     = "\033[34m"
	ansiCyan     = "\033[36m"
	ansiGray     = "\033[90m"
	ansiBoldCyan = "\033[1;36m"
)

// isTerminal reports whether the file descriptor is attached to a TTY.
// Used to auto-disable colors when piped to a file or another program.
func isTerminal(f *os.File) bool {
	fi, err := f.Stat()
	if err != nil {
		return false
	}
	return (fi.Mode() & os.ModeCharDevice) != 0
}

// noColor is true when NO_COLOR env var is set or stdout is not a TTY.
var noColor = os.Getenv("NO_COLOR") != "" || !isTerminal(os.Stdout)

// Quiet suppresses all non-error stdout when set true (used by --quiet flag).
// Errors and JSON results are NOT suppressed — they remain pipe-friendly.
var Quiet bool

// colorize wraps msg with ANSI codes when color is enabled, otherwise returns msg unchanged.
func colorize(code, msg string) string {
	if noColor {
		return msg
	}
	return code + msg + ansiReset
}

// Success prints a green check-mark message to stdout. Suppressed under --quiet.
func Success(msg string) {
	if Quiet {
		return
	}
	fmt.Println(colorize(ansiGreen, "[ok] "+msg))
}

// Error prints a red error message to stderr. Never suppressed.
func Error(msg string) {
	fmt.Fprintln(os.Stderr, colorize(ansiRed, "[err] "+msg))
}

// Warn prints a yellow warning message to stderr. Never suppressed.
func Warn(msg string) {
	fmt.Fprintln(os.Stderr, colorize(ansiYellow, "[warn] "+msg))
}

// Info prints a blue informational message to stdout. Suppressed under --quiet.
func Info(msg string) {
	if Quiet {
		return
	}
	fmt.Println(colorize(ansiBlue, "[i] "+msg))
}

// Bold prints a bold message to stdout. Suppressed under --quiet.
func Bold(msg string) {
	if Quiet {
		return
	}
	fmt.Println(colorize(ansiBold, msg))
}

// Gray prints a dim message to stdout (for hints, secondary info). Suppressed under --quiet.
func Gray(msg string) {
	if Quiet {
		return
	}
	fmt.Println(colorize(ansiGray, msg))
}

// FormatCyan returns a cyan-colored version of s.
func FormatCyan(s string) string { return colorize(ansiCyan, s) }

// FormatCyanBold returns a bold cyan-colored version of s.
func FormatCyanBold(s string) string { return colorize(ansiBoldCyan, s) }

// FormatGray returns a gray-colored version of s.
func FormatGray(s string) string { return colorize(ansiGray, s) }

// FormatGreen returns a green-colored version of s.
func FormatGreen(s string) string { return colorize(ansiGreen, s) }

// FormatRed returns a red-colored version of s.
func FormatRed(s string) string { return colorize(ansiRed, s) }

// FormatYellow returns a yellow-colored version of s.
func FormatYellow(s string) string { return colorize(ansiYellow, s) }
