package main

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/extrame/ole2"
)

const oleCbHdr = uint16(28)

type StyleHint struct {
	ObjectIndex                int      `json:"objectIndex"`
	Entry                      string   `json:"entry"`
	Hints                      []string `json:"hints,omitempty"`
	AsciiFunctionParen         int      `json:"asciiFunctionParen"`
	FullwidthTextParen         int      `json:"fullwidthTextParen"`
	LegacyEuclidOneParen       int      `json:"legacyEuclidOneParen"`
	LegacyEuclidTwoParen       int      `json:"legacyEuclidTwoParen"`
	ExplicitScriptFullSize     bool     `json:"explicitScriptFullSize"`
	ExplicitFractionFullSize   bool     `json:"explicitFractionFullSize"`
	ExplicitTopFullSize        bool     `json:"explicitTopFullSize"`
	ForceExplicitFenceTemplate bool     `json:"forceExplicitFenceTemplate"`
	ExplicitBlackColor         bool     `json:"explicitBlackColor"`
	FlatParenTemplate          bool     `json:"flatParenTemplate"`
	LetterGroupObarTemplate    bool     `json:"letterGroupObarTemplate"`
	TextFeComma                bool     `json:"textFeComma"`
	LegacyTextFeAscii          int      `json:"legacyTextFeAscii"`
	Error                      string   `json:"error,omitempty"`
}

func main() {
	if len(os.Args) != 3 {
		log.Fatalf("usage: extract_mtef_style_hints source.docx out.json")
	}
	hints, err := extract(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Dir(os.Args[2]), 0755); err != nil {
		log.Fatal(err)
	}
	data, _ := json.MarshalIndent(hints, "", "  ")
	if err := os.WriteFile(os.Args[2], append(data, '\n'), 0644); err != nil {
		log.Fatal(err)
	}
	fmt.Println(os.Args[2])
}

func extract(docxPath string) ([]StyleHint, error) {
	zr, err := zip.OpenReader(docxPath)
	if err != nil {
		return nil, err
	}
	defer zr.Close()
	entries := make([]string, 0)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "word/embeddings/") && strings.HasSuffix(strings.ToLower(f.Name), ".bin") {
			entries = append(entries, f.Name)
		}
	}
	sort.Slice(entries, func(i, j int) bool {
		return naturalLess(entries[i], entries[j])
	})
	out := make([]StyleHint, 0, len(entries))
	for _, entry := range entries {
		data, err := readZipEntry(&zr.Reader, entry)
		hint := StyleHint{ObjectIndex: objectIndex(entry), Entry: entry}
		if err != nil {
			hint.Error = err.Error()
			out = append(out, hint)
			continue
		}
		body, err := equationBody(data)
		if err != nil {
			hint.Error = err.Error()
			out = append(out, hint)
			continue
		}
		formula := formulaRegion(body)
		hint.AsciiFunctionParen = countCharMtcode(formula, 0x0028) + countCharMtcode(formula, 0x0029)
		hint.FullwidthTextParen = countCharMtcode(formula, 0xff08) + countCharMtcode(formula, 0xff09)
		hint.LegacyEuclidOneParen = countPlainChar(formula, 0x7f, 0x0028) + countPlainChar(formula, 0x7f, 0x0029)
		hint.LegacyEuclidTwoParen = countPlainChar(formula, 0x7e, 0x0028) + countPlainChar(formula, 0x7e, 0x0029)
		normalAsciiParen := hint.AsciiFunctionParen - hint.LegacyEuclidOneParen - hint.LegacyEuclidTwoParen
		hint.ExplicitScriptFullSize = bytes.Contains(formula, []byte{0x09, 0x65, 0x50, 0x01})
		hint.ExplicitFractionFullSize = hasExplicitFractionFullSize(formula)
		hint.ExplicitTopFullSize = hasExplicitTopFullSize(formula)
		hint.ForceExplicitFenceTemplate = hasFenceTemplate(formula)
		hint.ExplicitBlackColor = hasExplicitBlackColor(formula)
		hint.FlatParenTemplate = hasFlatParenTemplate(formula)
		hint.LetterGroupObarTemplate = hasObarTemplate(formula)
		hint.TextFeComma = bytes.Contains(formula, []byte{0x02, 0x00, 0x8c, 0x0c, 0xff})
		hint.LegacyTextFeAscii = countAsciiPlainTypeface(formula, 0x8c)
		if hint.LegacyEuclidOneParen > 0 && hint.LegacyEuclidTwoParen == 0 && normalAsciiParen == 0 {
			hint.Hints = append(hint.Hints, "asciiFlatParens")
		} else if hint.LegacyEuclidOneParen > 0 && hint.LegacyEuclidTwoParen > 0 && normalAsciiParen == 0 {
			hint.Hints = append(hint.Hints, "asciiFlatParens")
		} else if normalAsciiParen > 0 && hint.FullwidthTextParen > 0 {
			hint.Hints = append(hint.Hints, "mixedAsciiFullwidthParens")
		} else if hint.AsciiFunctionParen > hint.FullwidthTextParen {
			hint.Hints = append(hint.Hints, "asciiFlatParens")
		}
		if hint.FullwidthTextParen > hint.AsciiFunctionParen {
			hint.Hints = append(hint.Hints, "fullwidthTextParen")
		}
		if hint.ExplicitScriptFullSize {
			hint.Hints = append(hint.Hints, "explicitScriptFullSize")
		}
		if hint.ExplicitFractionFullSize {
			hint.Hints = append(hint.Hints, "explicitFractionFullSize")
		}
		if hint.ExplicitTopFullSize {
			hint.Hints = append(hint.Hints, "explicitTopFullSize")
		}
		if hint.ForceExplicitFenceTemplate {
			hint.Hints = append(hint.Hints, "forceExplicitFenceTemplate")
		}
		if hint.ExplicitBlackColor {
			hint.Hints = append(hint.Hints, "explicitBlackColor")
		}
		if hint.FlatParenTemplate {
			hint.Hints = append(hint.Hints, "flatParenTemplate")
		}
		if hint.LetterGroupObarTemplate {
			hint.Hints = append(hint.Hints, "letterGroupObarTemplate")
		}
		if hint.TextFeComma {
			hint.Hints = append(hint.Hints, "textFeComma")
		}
		if hint.FullwidthTextParen > 0 && hint.LegacyTextFeAscii > 0 {
			hint.Hints = append(hint.Hints, "legacyTextFeParenContent")
		}
		out = append(out, hint)
	}
	return out, nil
}

func hasFenceTemplate(body []byte) bool {
	return bytes.Contains(body, []byte{0x03, 0x00, 0x01, 0x03, 0x00}) ||
		bytes.Contains(body, []byte{0x03, 0x00, 0x03, 0x03, 0x00})
}

func hasFlatParenTemplate(body []byte) bool {
	return bytes.Contains(body, []byte{0x03, 0x00, 0x01, 0x00, 0x04})
}

func hasObarTemplate(body []byte) bool {
	return bytes.Contains(body, []byte{0x03, 0x00, 0x0d, 0x00, 0x00})
}

func formulaRegion(body []byte) []byte {
	start := 5
	for start < len(body) && body[start] != 0 {
		start++
	}
	start = min(len(body), start+2) // application-key terminator and equation options
	prefs := bytes.IndexByte(body[start:], 0x12)
	if prefs < 0 {
		return body[start:]
	}
	cursor := start + prefs + 1
	if cursor >= len(body) {
		return nil
	}
	cursor++ // EQN_PREFS options
	cursor = skipDimensionArray(body, cursor)
	cursor = skipDimensionArray(body, cursor)
	if cursor >= len(body) {
		return nil
	}
	styleCount := int(body[cursor])
	cursor++
	for i := 0; i < styleCount && cursor < len(body); i++ {
		fontDefinition := body[cursor]
		cursor++
		if fontDefinition != 0 && cursor < len(body) {
			cursor++
		}
	}
	return body[min(cursor, len(body)):]
}

func skipDimensionArray(body []byte, offset int) int {
	if offset >= len(body) {
		return len(body)
	}
	count := int(body[offset])
	byteOffset := offset + 1
	nibbleOffset := 0
	completed := 0
	for byteOffset+nibbleOffset/2 < len(body) && completed < count {
		packed := body[byteOffset+nibbleOffset/2]
		nibble := packed >> 4
		if nibbleOffset&1 != 0 {
			nibble = packed & 0x0f
		}
		nibbleOffset++
		if nibble == 0x0f {
			completed++
		}
	}
	return min(len(body), byteOffset+(nibbleOffset+1)/2)
}

func formulaTail(body []byte) []byte {
	lastLine := -1
	for i := 0; i+1 < len(body); i++ {
		if body[i] == 0x01 && (body[i+1] == 0x00 || body[i+1] == 0x01 || body[i+1] == 0x04) {
			lastLine = i
		}
	}
	if lastLine < 0 {
		return body
	}
	return body[lastLine:]
}

func hasExplicitTopFullSize(tail []byte) bool {
	return bytes.HasPrefix(tail, []byte{0x01, 0x00, 0x09, 0x65, 0x50, 0x01})
}

func hasExplicitBlackColor(tail []byte) bool {
	return bytes.HasPrefix(tail, []byte{
		0x01, 0x00, 0x10, 0x04, 0x00, 0x00, 0x00, 0x00,
		0x00, 0x00, 0x42, 0x6c, 0x61, 0x63, 0x6b, 0x00, 0x0f, 0x01,
	})
}

func hasExplicitFractionFullSize(body []byte) bool {
	header := []byte{0x03, 0x00, 0x0b, 0x00, 0x00}
	size := []byte{0x09, 0x65, 0x50, 0x01}
	for i := 0; i+len(header) <= len(body); i++ {
		if !bytes.Equal(body[i:i+len(header)], header) {
			continue
		}
		windowEnd := min(len(body), i+80)
		if bytes.Contains(body[i:windowEnd], size) {
			return true
		}
	}
	return false
}

func equationBody(oleData []byte) ([]byte, error) {
	ole, err := ole2.Open(bytes.NewReader(oleData), "")
	if err != nil {
		return nil, fmt.Errorf("open ole: %w", err)
	}
	dir, err := ole.ListDir()
	if err != nil {
		return nil, fmt.Errorf("list ole: %w", err)
	}
	for _, file := range dir {
		if file.Name() != "Equation Native" {
			continue
		}
		reader := ole.OpenFile(file, dir[0])
		native, err := io.ReadAll(reader)
		if err != nil {
			return nil, fmt.Errorf("read native: %w", err)
		}
		if len(native) < int(oleCbHdr) {
			return nil, fmt.Errorf("native too small")
		}
		cbHdr := binary.LittleEndian.Uint16(native[0:2])
		cbSize := binary.LittleEndian.Uint32(native[8:12])
		if cbHdr == 0 || int(cbHdr) > len(native) {
			return nil, fmt.Errorf("invalid cbHdr=%d", cbHdr)
		}
		end := min(len(native), int(cbHdr)+int(cbSize))
		return native[cbHdr:end], nil
	}
	return nil, fmt.Errorf("Equation Native stream not found")
}

func readZipEntry(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name != name {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		defer rc.Close()
		return io.ReadAll(rc)
	}
	return nil, fmt.Errorf("zip entry not found: %s", name)
}

func countSeq(data, seq []byte) int {
	count := 0
	for i := 0; i+len(seq) <= len(data); i++ {
		if bytes.Equal(data[i:i+len(seq)], seq) {
			count++
		}
	}
	return count
}

func countCharMtcode(data []byte, mtcode uint16) int {
	count := 0
	for i := 0; i+4 < len(data); i++ {
		if data[i] != 0x02 || data[i+1]&0x20 != 0 {
			continue
		}
		if binary.LittleEndian.Uint16(data[i+3:i+5]) == mtcode {
			count++
		}
	}
	return count
}

func countPlainChar(data []byte, typeface byte, mtcode uint16) int {
	return countSeq(data, []byte{0x02, 0x00, typeface, byte(mtcode), byte(mtcode >> 8)})
}

func countAsciiPlainTypeface(data []byte, typeface byte) int {
	count := 0
	for i := 0; i+4 < len(data); i++ {
		if data[i] == 0x02 && data[i+1] == 0x00 && data[i+2] == typeface && data[i+4] == 0x00 {
			count++
		}
	}
	return count
}

func objectIndex(source string) int {
	match := regexp.MustCompile(`oleObject(\d+)\.bin`).FindStringSubmatch(source)
	if len(match) != 2 {
		return 0
	}
	value, _ := strconv.Atoi(match[1])
	return value
}

func naturalLess(a, b string) bool {
	ra := regexp.MustCompile(`\d+|\D+`).FindAllString(a, -1)
	rb := regexp.MustCompile(`\d+|\D+`).FindAllString(b, -1)
	for i := 0; i < len(ra) && i < len(rb); i++ {
		ai, aerr := strconv.Atoi(ra[i])
		bi, berr := strconv.Atoi(rb[i])
		if aerr == nil && berr == nil {
			if ai != bi {
				return ai < bi
			}
			continue
		}
		if ra[i] != rb[i] {
			return ra[i] < rb[i]
		}
	}
	return len(ra) < len(rb)
}
