package output

import (
	"fmt"
	"os"
	"regexp"
	"strings"
	"unicode"

	"golang.org/x/term"
)

var ansiEscapeRe = regexp.MustCompile(`\x1b\[[0-9;]*[a-zA-Z]`)

// termWidth returns the terminal width in columns, defaulting to 120 when unknown.
func termWidth() int {
	if w, _, err := term.GetSize(int(os.Stdout.Fd())); err == nil && w > 0 {
		return w
	}
	return 120
}

// runeWidth returns the display width of s, counting CJK as 2 cols and ASCII as 1.
func runeWidth(s string) int {
	width := 0
	for _, r := range s {
		if isCJK(r) {
			width += 2
		} else {
			width++
		}
	}
	return width
}

// isCJK reports whether r is a CJK / wide character.
func isCJK(r rune) bool {
	if unicode.In(r, unicode.Han, unicode.Hiragana, unicode.Katakana, unicode.Hangul) {
		return true
	}
	return (r >= 0x2E80 && r <= 0x2EFF) ||
		(r >= 0x2F00 && r <= 0x2FDF) ||
		(r >= 0x3000 && r <= 0x303F) ||
		(r >= 0x3200 && r <= 0x32FF) ||
		(r >= 0x3300 && r <= 0x33FF) ||
		(r >= 0xF900 && r <= 0xFAFF) ||
		(r >= 0xFE30 && r <= 0xFE4F) ||
		(r >= 0xFF01 && r <= 0xFF60) ||
		(r >= 0xFFE0 && r <= 0xFFE6) ||
		(r >= 0x20000 && r <= 0x2A6DF)
}

func stripAnsi(s string) string {
	return ansiEscapeRe.ReplaceAllString(s, "")
}

// truncate cuts s to the given visible width, appending an ellipsis when overflowing.
func truncate(s string, maxWidth int) string {
	if maxWidth <= 0 {
		return ""
	}
	plain := stripAnsi(s)
	if runeWidth(plain) <= maxWidth {
		return s
	}
	width := 0
	var buf strings.Builder
	for _, r := range plain {
		rw := 1
		if isCJK(r) {
			rw = 2
		}
		if width+rw > maxWidth-1 {
			break
		}
		buf.WriteRune(r)
		width += rw
	}
	buf.WriteRune('\u2026')
	return buf.String()
}

// Table prints a bordered, colored table with the given headers and rows.
// Suppressed under --quiet. Designed for human reading; AI Agents use --json instead.
func Table(headers []string, rows [][]string) {
	if Quiet || len(headers) == 0 {
		return
	}

	cols := len(headers)
	colWidths := make([]int, cols)
	for i, h := range headers {
		w := runeWidth(strings.ToUpper(h))
		if w > colWidths[i] {
			colWidths[i] = w
		}
	}
	for _, row := range rows {
		for i := 0; i < cols && i < len(row); i++ {
			w := runeWidth(stripAnsi(row[i]))
			if w > colWidths[i] {
				colWidths[i] = w
			}
		}
	}

	totalWidth := 2 + cols
	for _, w := range colWidths {
		totalWidth += w
	}

	tw := termWidth()
	for totalWidth > tw && tw > 0 {
		maxIdx := 0
		for i := 1; i < cols; i++ {
			if colWidths[i] > colWidths[maxIdx] {
				maxIdx = i
			}
		}
		if colWidths[maxIdx] <= 3 {
			break
		}
		colWidths[maxIdx]--
		totalWidth--
	}

	buildHLine := func(left, mid, right, fill string) string {
		var sb strings.Builder
		sb.WriteString(left)
		for i, w := range colWidths {
			sb.WriteString(strings.Repeat(fill, w+2))
			if i < cols-1 {
				sb.WriteString(mid)
			}
		}
		sb.WriteString(right)
		return sb.String()
	}

	topLine := buildHLine("┌", "┬", "┐", "─")
	midLine := buildHLine("├", "┼", "┤", "─")
	botLine := buildHLine("└", "┴", "┘", "─")

	buildRow := func(cells []string, isHeader bool) string {
		var sb strings.Builder
		sb.WriteString("│")
		for i := 0; i < cols; i++ {
			cell := ""
			if i < len(cells) {
				cell = cells[i]
			}
			plain := stripAnsi(cell)
			if runeWidth(plain) > colWidths[i] {
				cell = truncate(plain, colWidths[i])
				plain = stripAnsi(cell)
			}
			if isHeader {
				cell = FormatCyanBold(strings.ToUpper(plain))
			}
			displayWidth := runeWidth(stripAnsi(cell))
			padding := colWidths[i] - displayWidth
			if padding < 0 {
				padding = 0
			}
			sb.WriteString(" ")
			sb.WriteString(cell)
			sb.WriteString(strings.Repeat(" ", padding))
			sb.WriteString(" │")
		}
		return sb.String()
	}

	fmt.Println(topLine)
	fmt.Println(buildRow(headers, true))
	fmt.Println(midLine)
	for _, row := range rows {
		fmt.Println(buildRow(row, false))
	}
	fmt.Println(botLine)
}
