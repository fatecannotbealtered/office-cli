package common

import (
	"bufio"
	"bytes"
	"io"
	"unicode/utf8"
)

// UTF8BOM is the byte order mark for UTF-8. Excel on Windows expects this when
// opening CSV files, otherwise CJK characters render as garbage (mojibake).
var UTF8BOM = []byte{0xEF, 0xBB, 0xBF}

// StripBOM returns data without a leading UTF-8 / UTF-16 BOM. office-cli reads
// XML payloads from docx/pptx where some authoring tools insert BOMs that break
// strict parsers — strip defensively.
func StripBOM(data []byte) []byte {
	switch {
	case len(data) >= 3 && bytes.Equal(data[:3], UTF8BOM):
		return data[3:]
	case len(data) >= 2 && data[0] == 0xFE && data[1] == 0xFF:
		return data[2:]
	case len(data) >= 2 && data[0] == 0xFF && data[1] == 0xFE:
		return data[2:]
	}
	return data
}

// EnsureUTF8 returns data unchanged when it is already valid UTF-8.
// Otherwise it returns a best-effort UTF-8 representation (each invalid byte is
// replaced by U+FFFD). office-cli uses this for human-readable text output;
// authoring formats (docx/pptx) are guaranteed UTF-8 by the spec and are never
// rewritten through this helper.
func EnsureUTF8(data []byte) []byte {
	if utf8.Valid(data) {
		return data
	}
	var out bytes.Buffer
	out.Grow(len(data))
	for i := 0; i < len(data); {
		r, size := utf8.DecodeRune(data[i:])
		if r == utf8.RuneError && size == 1 {
			out.WriteRune('\uFFFD')
			i++
			continue
		}
		out.WriteRune(r)
		i += size
	}
	return out.Bytes()
}

// WriteUTF8 wraps an io.Writer with UTF-8 line buffering and an optional BOM
// prefix. Used for CSV / text exports.
func WriteUTF8(w io.Writer, withBOM bool) (*bufio.Writer, error) {
	bw := bufio.NewWriter(w)
	if withBOM {
		if _, err := bw.Write(UTF8BOM); err != nil {
			return nil, err
		}
	}
	return bw, nil
}
