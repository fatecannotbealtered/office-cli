package output

import (
	"bytes"
	"encoding/json"
	"os"
	"strings"
	"testing"
)

func TestPrintJSON(t *testing.T) {
	old := os.Stdout
	r, w, _ := os.Pipe()
	os.Stdout = w

	PrintJSON(map[string]any{"key": "value", "num": 42})

	_ = w.Close()
	os.Stdout = old

	var buf bytes.Buffer
	_, _ = buf.ReadFrom(r)

	var m map[string]any
	if err := json.Unmarshal(buf.Bytes(), &m); err != nil {
		t.Fatalf("invalid JSON: %v\nraw: %s", err, buf.String())
	}
	if m["key"] != "value" {
		t.Errorf("key = %v, want value", m["key"])
	}
}

func TestPrintErrorJSON(t *testing.T) {
	old := os.Stderr
	r, w, _ := os.Pipe()
	os.Stderr = w

	PrintErrorJSON("test error", ErrValidation)

	_ = w.Close()
	os.Stderr = old

	var buf bytes.Buffer
	_, _ = buf.ReadFrom(r)

	var m map[string]any
	if err := json.Unmarshal(buf.Bytes(), &m); err != nil {
		t.Fatalf("invalid JSON: %v\nraw: %s", err, buf.String())
	}
	if m["error"] != "test error" {
		t.Errorf("error = %v", m["error"])
	}
	if m["errorCode"] != string(ErrValidation) {
		t.Errorf("errorCode = %v", m["errorCode"])
	}
	if m["hint"] == nil || m["hint"] == "" {
		t.Error("hint should be present")
	}
}

func TestPrintErrorJSONWithFile(t *testing.T) {
	old := os.Stderr
	r, w, _ := os.Pipe()
	os.Stderr = w

	PrintErrorJSONWithFile("not found", ErrFileNotFound, "test.xlsx")

	_ = w.Close()
	os.Stderr = old

	var buf bytes.Buffer
	_, _ = buf.ReadFrom(r)

	var m map[string]any
	if err := json.Unmarshal(buf.Bytes(), &m); err != nil {
		t.Fatalf("invalid JSON: %v", err)
	}
	if m["file"] != "test.xlsx" {
		t.Errorf("file = %v", m["file"])
	}
}

func TestHintForErrorCode(t *testing.T) {
	codes := []ErrorCode{
		ErrFileNotFound, ErrInvalidFormat, ErrCorrupted, ErrPermission,
		ErrValidation, ErrNotFound, ErrEngine, ErrToolMissing, ErrUnknown,
	}
	for _, code := range codes {
		hint := HintForErrorCode(code)
		if hint == "" {
			t.Errorf("HintForErrorCode(%q) returned empty", code)
		}
	}
}

func TestErrorCodes(t *testing.T) {
	// Verify error codes match spec
	if ErrFileNotFound != "FILE_NOT_FOUND" {
		t.Errorf("ErrFileNotFound = %q", ErrFileNotFound)
	}
	if ErrValidation != "VALIDATION_ERROR" {
		t.Errorf("ErrValidation = %q", ErrValidation)
	}
	if ErrPermission != "PERMISSION_DENIED" {
		t.Errorf("ErrPermission = %q", ErrPermission)
	}
}

func TestFilterMap(t *testing.T) {
	m := map[string]any{
		"name":  "Alice",
		"age":   30,
		"email": "a@b.com",
	}

	// No filter — return all
	all := FilterMap(m, nil)
	if len(all) != 3 {
		t.Errorf("unfiltered: got %d keys, want 3", len(all))
	}

	// Filter to specific fields
	filtered := FilterMap(m, []string{"name", "email"})
	if len(filtered) != 2 {
		t.Errorf("filtered: got %d keys, want 2", len(filtered))
	}
	if filtered["name"] != "Alice" {
		t.Errorf("name = %v", filtered["name"])
	}

	// Case insensitive
	filtered2 := FilterMap(m, []string{"NAME"})
	if filtered2["name"] != "Alice" {
		t.Errorf("case insensitive: name = %v", filtered2["name"])
	}

	// Missing field
	filtered3 := FilterMap(m, []string{"nonexistent"})
	if len(filtered3) != 0 {
		t.Errorf("missing field: got %d keys, want 0", len(filtered3))
	}
}

func TestColorDisabled(t *testing.T) {
	// When noColor is true, colorize should return raw text
	old := noColor
	noColor = true
	defer func() { noColor = old }()

	got := colorize(ansiRed, "hello")
	if got != "hello" {
		t.Errorf("colorize with noColor = %q, want %q", got, "hello")
	}
}

func TestColorEnabled(t *testing.T) {
	old := noColor
	noColor = false
	defer func() { noColor = old }()

	got := colorize(ansiRed, "hello")
	if !strings.Contains(got, ansiRed) {
		t.Error("colorize should contain ANSI code when color enabled")
	}
	if !strings.Contains(got, "hello") {
		t.Error("colorize should contain the message")
	}
}

func TestQuietSuppression(t *testing.T) {
	old := Quiet
	Quiet = true
	defer func() { Quiet = old }()

	oldStdout := os.Stdout
	r, w, _ := os.Pipe()
	os.Stdout = w

	Success("test")
	Info("test")
	Bold("test")
	Gray("test")

	_ = w.Close()
	os.Stdout = oldStdout

	var buf bytes.Buffer
	_, _ = buf.ReadFrom(r)
	if buf.Len() != 0 {
		t.Errorf("Quiet mode should suppress output, got %q", buf.String())
	}
}

func TestQuietDoesNotSuppressError(t *testing.T) {
	old := Quiet
	Quiet = true
	defer func() { Quiet = old }()

	oldStderr := os.Stderr
	r, w, _ := os.Pipe()
	os.Stderr = w

	Error("test error")

	_ = w.Close()
	os.Stderr = oldStderr

	var buf bytes.Buffer
	_, _ = buf.ReadFrom(r)
	if buf.Len() == 0 {
		t.Error("Quiet mode should NOT suppress errors")
	}
}

func TestFlatTypes(t *testing.T) {
	// Verify flat types have correct JSON tags
	cell := FlatCell{Sheet: "S", Ref: "A1", Value: "v"}
	data, err := json.Marshal(cell)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), `"sheet":"S"`) {
		t.Errorf("unexpected JSON: %s", data)
	}

	// omitempty on Value
	cell2 := FlatCell{Sheet: "S", Ref: "A1"}
	data2, _ := json.Marshal(cell2)
	if strings.Contains(string(data2), "value") {
		t.Errorf("empty value should be omitted: %s", data2)
	}
}
