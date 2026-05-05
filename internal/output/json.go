package output

import (
	"encoding/json"
	"fmt"
	"os"
)

// PrintJSON outputs v as indented JSON to stdout. Used for --json results.
func PrintJSON(v any) {
	data, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		fmt.Fprintf(os.Stderr, "json marshal error: %v\n", err)
		return
	}
	fmt.Println(string(data))
}

// ErrorCode is a stable, machine-readable classification of failure cases.
// AI Agents should switch on these codes (rather than parse messages) when
// deciding whether to retry, prompt the user, or abort.
type ErrorCode string

const (
	// ErrFileNotFound indicates a referenced input file does not exist.
	ErrFileNotFound ErrorCode = "FILE_NOT_FOUND"
	// ErrInvalidFormat indicates the file extension/content does not match the expected format.
	ErrInvalidFormat ErrorCode = "INVALID_FORMAT"
	// ErrCorrupted indicates a structurally corrupt or password-protected document.
	ErrCorrupted ErrorCode = "CORRUPTED_FILE"
	// ErrPermission indicates an OS-level permission issue (read/write denied).
	ErrPermission ErrorCode = "PERMISSION_DENIED"
	// ErrValidation indicates bad command arguments or invalid spec content.
	ErrValidation ErrorCode = "VALIDATION_ERROR"
	// ErrNotFound indicates a target inside a document (sheet, page, slide) does not exist.
	ErrNotFound ErrorCode = "NOT_FOUND"
	// ErrEngine indicates a low-level error from a third-party document engine.
	ErrEngine ErrorCode = "ENGINE_ERROR"
	// ErrToolMissing indicates a required external tool (pandoc, libreoffice) is not available.
	ErrToolMissing ErrorCode = "TOOL_MISSING"
	// ErrUnknown is the fallback bucket.
	ErrUnknown ErrorCode = "UNKNOWN_ERROR"
)

// HintForErrorCode returns an actionable hint for the given error code,
// shown in JSON error responses to help AI Agents take corrective action.
func HintForErrorCode(code ErrorCode) string {
	switch code {
	case ErrFileNotFound:
		return "Verify the file path exists and is reachable from the current working directory"
	case ErrInvalidFormat:
		return "The file extension or content does not match the expected format (e.g. .xlsx, .docx, .pdf)"
	case ErrCorrupted:
		return "The document is corrupted or password-protected; try opening it manually first"
	case ErrPermission:
		return "Check OS file permissions; ensure the file is not opened/locked by another program"
	case ErrValidation:
		return "Check command arguments, flag values and JSON spec structure"
	case ErrNotFound:
		return "The requested sheet/page/slide does not exist in this document"
	case ErrEngine:
		return "An internal engine error occurred; re-run with --json to see details"
	case ErrToolMissing:
		return "An optional external tool is missing; run 'office-cli doctor' to see installation hints"
	case ErrUnknown:
		return "An unexpected error occurred; re-run with --json for details and check the input file"
	default:
		return ""
	}
}

// jsonError is the wire shape of every JSON error response.
type jsonError struct {
	Error     string    `json:"error"`
	ErrorCode ErrorCode `json:"errorCode"`
	Hint      string    `json:"hint,omitempty"`
	File      string    `json:"file,omitempty"`
}

// PrintErrorJSON writes a structured error to stderr.
func PrintErrorJSON(msg string, code ErrorCode) {
	PrintErrorJSONWithFile(msg, code, "")
}

// PrintErrorJSONWithFile writes a structured error to stderr including the offending file path.
// An optional hint override can be provided; if empty, the default hint for the error code is used.
func PrintErrorJSONWithFile(msg string, code ErrorCode, file string, hintOverride ...string) {
	hint := HintForErrorCode(code)
	if len(hintOverride) > 0 && hintOverride[0] != "" {
		hint = hintOverride[0]
	}
	payload := jsonError{
		Error:     msg,
		ErrorCode: code,
		Hint:      hint,
		File:      file,
	}
	data, err := json.MarshalIndent(payload, "", "  ")
	if err != nil {
		fmt.Fprintf(os.Stderr, `{"error": %q, "errorCode": %q}`+"\n", msg, code)
		return
	}
	fmt.Fprintln(os.Stderr, string(data))
}
