package common

import (
	"bytes"
	"testing"
)

func TestStripBOM(t *testing.T) {
	withBOM := append([]byte{0xEF, 0xBB, 0xBF}, []byte("hello")...)
	out := StripBOM(withBOM)
	if string(out) != "hello" {
		t.Errorf("expected hello, got %q", string(out))
	}
	if !bytes.Equal(StripBOM([]byte("plain")), []byte("plain")) {
		t.Errorf("plain should be unchanged")
	}
	if !bytes.Equal(StripBOM(nil), nil) {
		t.Errorf("nil should stay nil")
	}
}

func TestEnsureUTF8(t *testing.T) {
	out := EnsureUTF8([]byte("hello 世界"))
	if string(out) != "hello 世界" {
		t.Errorf("expected utf-8 unchanged, got %q", string(out))
	}
	bad := []byte{0xFF, 0xFE, 0x68, 0x69}
	clean := EnsureUTF8(bad)
	if !bytes.ContainsRune(clean, 'h') {
		t.Errorf("expected ASCII to survive sanitization")
	}
}

func TestUTF8BOM(t *testing.T) {
	if len(UTF8BOM) != 3 || UTF8BOM[0] != 0xEF || UTF8BOM[1] != 0xBB || UTF8BOM[2] != 0xBF {
		t.Errorf("UTF8BOM constant is wrong: %v", UTF8BOM)
	}
}
