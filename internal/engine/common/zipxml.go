// Package common contains helpers shared between the docx (Word) and pptx (PowerPoint)
// engines. Both formats are ZIP archives containing XML parts; this package gives
// us a single place to read entries, rewrite entries, and parse a few cross-cutting
// metadata structures (core.xml, app.xml).
package common

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
)

// CoreProps mirrors the Dublin-Core metadata stored at docProps/core.xml.
type CoreProps struct {
	XMLName        xml.Name `xml:"coreProperties"`
	Title          string   `xml:"title"`
	Subject        string   `xml:"subject"`
	Creator        string   `xml:"creator"`
	Keywords       string   `xml:"keywords"`
	Description    string   `xml:"description"`
	LastModifiedBy string   `xml:"lastModifiedBy"`
	Revision       string   `xml:"revision"`
	Created        string   `xml:"created"`
	Modified       string   `xml:"modified"`
	Category       string   `xml:"category"`
}

// AppProps mirrors the Office-specific metadata at docProps/app.xml.
type AppProps struct {
	XMLName     xml.Name `xml:"Properties"`
	Application string   `xml:"Application"`
	AppVersion  string   `xml:"AppVersion"`
	Pages       int      `xml:"Pages"`
	Words       int      `xml:"Words"`
	Characters  int      `xml:"Characters"`
	Slides      int      `xml:"Slides"`
	Company     string   `xml:"Company"`
}

// ReadEntry returns the bytes of one zip entry. Returns os.ErrNotExist if absent.
func ReadEntry(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, fmt.Errorf("opening %s: %w", name, err)
			}
			defer func() { _ = rc.Close() }()
			data, err := io.ReadAll(rc)
			if err != nil {
				return nil, fmt.Errorf("reading %s: %w", name, err)
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("entry not found: %s: %w", name, os.ErrNotExist)
}

// OpenReader opens path as a zip archive. Caller must close the returned reader.
func OpenReader(path string) (*zip.ReadCloser, error) {
	return zip.OpenReader(path)
}

// ReadCoreProps unmarshals docProps/core.xml. Missing parts are silently empty.
func ReadCoreProps(zr *zip.Reader) CoreProps {
	var c CoreProps
	data, err := ReadEntry(zr, "docProps/core.xml")
	if err != nil {
		return c
	}
	_ = xml.Unmarshal(data, &c)
	return c
}

// ReadAppProps unmarshals docProps/app.xml. Missing parts are silently empty.
func ReadAppProps(zr *zip.Reader) AppProps {
	var a AppProps
	data, err := ReadEntry(zr, "docProps/app.xml")
	if err != nil {
		return a
	}
	_ = xml.Unmarshal(data, &a)
	return a
}

// RewriteEntries copies path to outPath, replacing the entries in `replacements` with
// the new bytes. Other entries are passed through unchanged. Comments / images /
// styles etc. survive intact, which is critical for editing real-world documents.
//
// When path == outPath (in-place editing), the source is read into memory first
// so the output file can safely overwrite the source without corruption.
func RewriteEntries(path, outPath string, replacements map[string][]byte) error {
	if len(replacements) == 0 {
		return errors.New("no replacements provided")
	}

	// Read the entire source into memory so we can safely overwrite it.
	srcData, err := os.ReadFile(path)
	if err != nil {
		return err
	}
	zr, err := zip.NewReader(bytes.NewReader(srcData), int64(len(srcData)))
	if err != nil {
		return err
	}

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}

	zw := zip.NewWriter(out)

	written := map[string]bool{}
	for _, f := range zr.File {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: f.Name, Method: f.Method})
		if err != nil {
			_ = zw.Close()
			_ = out.Close()
			return err
		}
		if newBytes, ok := replacements[f.Name]; ok {
			if _, err := w.Write(newBytes); err != nil {
				_ = zw.Close()
				_ = out.Close()
				return err
			}
			written[f.Name] = true
			continue
		}
		rc, err := f.Open()
		if err != nil {
			_ = zw.Close()
			_ = out.Close()
			return err
		}
		if _, err := io.Copy(w, rc); err != nil {
			_ = rc.Close()
			_ = zw.Close()
			_ = out.Close()
			return err
		}
		_ = rc.Close()
		written[f.Name] = true
	}

	// Add any new entries from replacements that don't exist in the source.
	for name, newBytes := range replacements {
		if written[name] {
			continue
		}
		w, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			_ = zw.Close()
			_ = out.Close()
			return err
		}
		if _, err := w.Write(newBytes); err != nil {
			_ = zw.Close()
			_ = out.Close()
			return err
		}
	}

	if err := zw.Close(); err != nil {
		_ = out.Close()
		return err
	}
	return out.Close()
}

// ReplaceInBytes returns a copy of in with every occurrence of `old` replaced by `new`.
// A convenience wrapper used by simple text-replacement workflows.
func ReplaceInBytes(in []byte, old, newVal string) []byte {
	if old == "" {
		return in
	}
	return bytes.ReplaceAll(in, []byte(old), []byte(newVal))
}

// WriteZipEntryToFile copies one zip entry's contents to disk at dst.
// The destination directory must already exist; callers are expected to call
// os.MkdirAll up-front (this lets the caller batch many writes into one mkdir).
func WriteZipEntryToFile(f *zip.File, dst string) error {
	rc, err := f.Open()
	if err != nil {
		return fmt.Errorf("opening %s: %w", f.Name, err)
	}
	defer func() { _ = rc.Close() }()

	out, err := os.Create(dst)
	if err != nil {
		return fmt.Errorf("creating %s: %w", dst, err)
	}
	defer func() { _ = out.Close() }()

	if _, err := io.Copy(out, rc); err != nil {
		return fmt.Errorf("writing %s: %w", dst, err)
	}
	return nil
}

// WriteNewZip creates a new zip file at path from the given string content map.
// Used by Word and PPT Create functions to build minimal OOXML archives.
func WriteNewZip(path string, files map[string]string) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer func() { _ = f.Close() }()

	w := zip.NewWriter(f)
	for name, content := range files {
		hdr := &zip.FileHeader{Name: name, Method: zip.Deflate}
		entry, err := w.CreateHeader(hdr)
		if err != nil {
			return err
		}
		if _, err := entry.Write([]byte(content)); err != nil {
			return err
		}
	}
	return w.Close()
}

// GuessMediaType returns the IANA media type for a filename based on its
// extension. Returns "" when the extension is unknown. We avoid mime.TypeByExtension
// to keep the result deterministic across operating systems.
func GuessMediaType(name string) string {
	ext := ""
	for i := len(name) - 1; i >= 0; i-- {
		if name[i] == '.' {
			ext = name[i:]
			break
		}
		if name[i] == '/' || name[i] == '\\' {
			break
		}
	}
	switch ext {
	case ".png":
		return "image/png"
	case ".jpg", ".jpeg":
		return "image/jpeg"
	case ".gif":
		return "image/gif"
	case ".bmp":
		return "image/bmp"
	case ".svg":
		return "image/svg+xml"
	case ".webp":
		return "image/webp"
	case ".tif", ".tiff":
		return "image/tiff"
	case ".emf":
		return "image/x-emf"
	case ".wmf":
		return "image/x-wmf"
	}
	return ""
}
