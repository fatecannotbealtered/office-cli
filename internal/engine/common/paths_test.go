package common

import (
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

func TestDetectFormat(t *testing.T) {
	cases := map[string]string{
		"a.docx":            "docx",
		"a.DOCX":            "docx",
		"a.xlsx":            "xlsx",
		"a.pptx":            "pptx",
		"a.pdf":             "pdf",
		"a.csv":             "csv",
		"a.txt":             "txt",
		"a.unknown":         "",
		"path/to/sample.md": "md",
	}
	for in, want := range cases {
		if got := DetectFormat(in); got != want {
			t.Errorf("DetectFormat(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestEnsureExtension(t *testing.T) {
	if err := EnsureExtension("foo.docx", "docx"); err != nil {
		t.Errorf("docx single: %v", err)
	}
	if err := EnsureExtension("foo.docx", "docx", "pdf"); err != nil {
		t.Errorf("docx multi: %v", err)
	}
	if err := EnsureExtension("foo.txt", "docx"); err == nil {
		t.Errorf("expected error for mismatch")
	}
}

func TestNormalizePath(t *testing.T) {
	got, err := NormalizePath("./foo.txt")
	if err != nil {
		t.Fatalf("normalize: %v", err)
	}
	if !filepath.IsAbs(got) {
		t.Errorf("expected absolute path, got %q", got)
	}

	if runtime.GOOS != "windows" {
		got, err = NormalizePath("~/")
		if err != nil {
			t.Fatalf("normalize tilde: %v", err)
		}
		if strings.HasPrefix(got, "~") {
			t.Errorf("tilde not expanded: %q", got)
		}
	}
}
