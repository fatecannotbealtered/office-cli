package cmd

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// TestGoFmt checks that every .go file in the repository is gofmt-clean.
// This integrates formatting validation into `go test ./...` so developers
// catch formatting issues locally without running a separate CI step.
func TestGoFmt(t *testing.T) {
	// Find the module root by looking for go.mod
	root := findModuleRoot(t)

	cmd := exec.Command("gofmt", "-l", ".")
	cmd.Dir = root
	out, err := cmd.Output()
	if err != nil {
		t.Fatalf("gofmt failed: %v", err)
	}

	files := strings.TrimSpace(string(out))
	if files == "" {
		return // all clean
	}

	// Filter out generated/vendor files
	var dirty []string
	for _, f := range strings.Split(files, "\n") {
		f = strings.TrimSpace(f)
		if f == "" {
			continue
		}
		// Skip tmp/ and demo/ directories (test fixtures)
		if strings.HasPrefix(f, "tmp/") || strings.HasPrefix(f, "demo/") {
			continue
		}
		dirty = append(dirty, f)
	}

	if len(dirty) > 0 {
		t.Errorf("the following files are not gofmt-formatted:\n%s\n\nRun: gofmt -w .", strings.Join(dirty, "\n"))
	}
}

// TestGoVet checks that `go vet ./...` passes.
func TestGoVet(t *testing.T) {
	root := findModuleRoot(t)

	cmd := exec.Command("go", "vet", "./...")
	cmd.Dir = root
	cmd.Env = append(os.Environ(), "GOFLAGS=-mod=mod")
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("go vet failed:\n%s\n%s", string(out), err)
	}
}

// findModuleRoot walks up from the test file's directory to find go.mod.
func findModuleRoot(t *testing.T) string {
	t.Helper()
	dir, err := os.Getwd()
	if err != nil {
		t.Fatalf("getting working directory: %v", err)
	}
	for {
		if _, err := os.Stat(filepath.Join(dir, "go.mod")); err == nil {
			return dir
		}
		parent := filepath.Dir(dir)
		if parent == dir {
			t.Fatal("could not find go.mod in any parent directory")
		}
		dir = parent
	}
}
