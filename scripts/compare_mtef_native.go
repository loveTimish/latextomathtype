package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/binary"
	"encoding/csv"
	"encoding/hex"
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

type OleObject struct {
	Index      int            `json:"index"`
	Entry      string         `json:"entry"`
	NativeSize int            `json:"nativeSize"`
	BodySize   int            `json:"bodySize"`
	BodySHA256 string         `json:"bodySha256"`
	HeaderHex  string         `json:"headerHex"`
	PrefixHex  string         `json:"prefixHex"`
	MtefVersion int           `json:"mtefVersion,omitempty"`
	MtefProduct int           `json:"mtefProduct,omitempty"`
	Records    map[string]int `json:"records"`
	TailSize   int            `json:"tailSize"`
	TailSHA256 string         `json:"tailSha256"`
	TailRecords map[string]int `json:"tailRecords"`
	Error      string         `json:"error,omitempty"`
	body       []byte
	content    []byte
	tail       []byte
	contentCandidates [][]byte
	tailCandidates    [][]byte
}

type PairRow struct {
	Index              int     `json:"index"`
	SourceIndex        int     `json:"sourceIndex"`
	GeneratedIndex     int     `json:"generatedIndex"`
	SourceEntry        string  `json:"sourceEntry"`
	GeneratedEntry     string  `json:"generatedEntry"`
	Latex              string  `json:"latex,omitempty"`
	SourceBodySize     int     `json:"sourceBodySize"`
	GeneratedBodySize  int     `json:"generatedBodySize"`
	BodySizeRatio      float64 `json:"bodySizeRatio,omitempty"`
	BodyHashEqual      bool    `json:"bodyHashEqual"`
	RecordCosine       float64 `json:"recordCosine"`
	TailSizeRatio      float64 `json:"tailSizeRatio,omitempty"`
	TailHashEqual      bool    `json:"tailHashEqual"`
	TailRecordCosine   float64 `json:"tailRecordCosine"`
	TailCoreRecordCosine float64 `json:"tailCoreRecordCosine"`
	CommonSuffixBytes  int     `json:"commonSuffixBytes"`
	CommonSuffixRatio  float64 `json:"commonSuffixRatio,omitempty"`
	AlignmentSuspect   bool    `json:"alignmentSuspect"`
	SourceError        string  `json:"sourceError,omitempty"`
	GeneratedError     string  `json:"generatedError,omitempty"`
	SourceRecordTotal  int     `json:"sourceRecordTotal"`
	GeneratedRecordTotal int   `json:"generatedRecordTotal"`
	SourceMtefVersion  int     `json:"sourceMtefVersion,omitempty"`
	SourceMtefProduct  int     `json:"sourceMtefProduct,omitempty"`
}

type Summary struct {
	Source              string    `json:"source"`
	Generated           string    `json:"generated"`
	SourceObjects       int       `json:"sourceObjects"`
	GeneratedObjects    int       `json:"generatedObjects"`
	PairedObjects       int       `json:"pairedObjects"`
	ReadableSource      int       `json:"readableSource"`
	ReadableGenerated   int       `json:"readableGenerated"`
	BodyHashEqual       int       `json:"bodyHashEqual"`
	MedianBodySizeRatio float64   `json:"medianBodySizeRatio,omitempty"`
	MedianRecordCosine  float64   `json:"medianRecordCosine,omitempty"`
	Pairs               []PairRow `json:"pairs"`
}

func main() {
	if len(os.Args) != 4 && len(os.Args) != 6 {
		log.Fatalf("usage: compare_mtef_native source.docx generated.docx out_dir [source.report.json generated.request.json]")
	}
	source := os.Args[1]
	generated := os.Args[2]
	outDir := os.Args[3]
	if err := os.MkdirAll(outDir, 0755); err != nil {
		log.Fatal(err)
	}
	var sourceReport, generatedRequest string
	if len(os.Args) == 6 {
		sourceReport = os.Args[4]
		generatedRequest = os.Args[5]
	}
	summary, err := compare(source, generated, sourceReport, generatedRequest)
	if err != nil {
		log.Fatal(err)
	}
	stem := strings.TrimSuffix(filepath.Base(source), filepath.Ext(source))
	jsonPath := filepath.Join(outDir, stem+"_mtef_summary.json")
	csvPath := filepath.Join(outDir, stem+"_mtef_pairs.csv")
	data, _ := json.MarshalIndent(summary, "", "  ")
	if err := os.WriteFile(jsonPath, append(data, '\n'), 0644); err != nil {
		log.Fatal(err)
	}
	if err := writeCSV(csvPath, summary.Pairs); err != nil {
		log.Fatal(err)
	}
	if err := writeWorstDumps(outDir, stem, source, generated, summary.Pairs); err != nil {
		log.Fatal(err)
	}
	fmt.Println(jsonPath)
}

func compare(sourcePath, generatedPath, sourceReportPath, generatedRequestPath string) (*Summary, error) {
	source, err := extractObjects(sourcePath)
	if err != nil {
		return nil, err
	}
	generated, err := extractObjects(generatedPath)
	if err != nil {
		return nil, err
	}
	indexPairs := ordinalPairs(len(source), len(generated))
	latexBySource := map[int]string{}
	if sourceReportPath != "" && generatedRequestPath != "" {
		sourceLatex, err := reportLatexSequence(sourceReportPath)
		if err != nil {
			return nil, err
		}
		generatedLatex, err := requestLatexSequence(generatedRequestPath)
		if err != nil {
			return nil, err
		}
		indexPairs = latexPairs(sourceLatex, generatedLatex, len(source), len(generated))
		for _, item := range sourceLatex {
			if item.ObjectIndex > 0 {
				latexBySource[item.ObjectIndex-1] = item.Latex
			}
		}
	}
	n := len(indexPairs)
	pairs := make([]PairRow, 0, n)
	ratios := make([]float64, 0, n)
	cosines := make([]float64, 0, n)
	hashEqual := 0
	for pairIndex, pair := range indexPairs {
		sourceIndex := pair[0]
		generatedIndex := pair[1]
		s := source[sourceIndex]
		g := generated[generatedIndex]
		row := PairRow{
			Index: pairIndex + 1,
			SourceIndex: sourceIndex + 1,
			GeneratedIndex: generatedIndex + 1,
			SourceEntry: s.Entry,
			GeneratedEntry: g.Entry,
			Latex: latexBySource[sourceIndex],
			SourceBodySize: s.BodySize,
			GeneratedBodySize: g.BodySize,
			BodyHashEqual: s.BodySHA256 != "" && s.BodySHA256 == g.BodySHA256,
			RecordCosine: bestCandidateCosine(s.contentCandidates, g.contentCandidates),
			TailHashEqual: s.TailSHA256 != "" && s.TailSHA256 == g.TailSHA256,
			TailRecordCosine: bestCandidateCosine(s.tailCandidates, g.tailCandidates),
			TailCoreRecordCosine: bestCandidateCoreCosine(s.tailCandidates, g.tailCandidates),
			CommonSuffixBytes: commonSuffixBytes(s.body, g.body),
			SourceError: s.Error,
			GeneratedError: g.Error,
			SourceRecordTotal: recordTotal(s.Records),
			GeneratedRecordTotal: recordTotal(g.Records),
			SourceMtefVersion: s.MtefVersion,
			SourceMtefProduct: s.MtefProduct,
		}
		if s.BodySize > 0 && g.BodySize > 0 {
			row.BodySizeRatio = float64(g.BodySize) / float64(s.BodySize)
			ratios = append(ratios, row.BodySizeRatio)
		}
		if s.TailSize > 0 && g.TailSize > 0 {
			row.TailSizeRatio = float64(g.TailSize) / float64(s.TailSize)
			extremeTailRatio := row.TailSizeRatio < 0.50 || row.TailSizeRatio > 2.00
			row.AlignmentSuspect = extremeTailRatio && row.RecordCosine < 0.90 && row.TailRecordCosine < 0.90
		}
		if s.BodySize > 0 && g.BodySize > 0 {
			row.CommonSuffixRatio = float64(row.CommonSuffixBytes) / float64(min(s.BodySize, g.BodySize))
		}
		if row.RecordCosine > 0 {
			cosines = append(cosines, row.RecordCosine)
		}
		if row.BodyHashEqual {
			hashEqual++
		}
		pairs = append(pairs, row)
	}
	return &Summary{
		Source: sourcePath,
		Generated: generatedPath,
		SourceObjects: len(source),
		GeneratedObjects: len(generated),
		PairedObjects: n,
		ReadableSource: countReadable(source),
		ReadableGenerated: countReadable(generated),
		BodyHashEqual: hashEqual,
		MedianBodySizeRatio: median(ratios),
		MedianRecordCosine: median(cosines),
		Pairs: pairs,
	}, nil
}

func extractObjects(docxPath string) ([]OleObject, error) {
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
	out := make([]OleObject, 0, len(entries))
	for i, entry := range entries {
		data, err := readZipEntry(&zr.Reader, entry)
		obj := OleObject{Index: i + 1, Entry: entry}
		if err != nil {
			obj.Error = err.Error()
			out = append(out, obj)
			continue
		}
		fillNativeStats(&obj, data)
		out = append(out, obj)
	}
	return out, nil
}

func fillNativeStats(obj *OleObject, oleData []byte) {
	ole, err := ole2.Open(bytes.NewReader(oleData), "")
	if err != nil {
		obj.Error = "open ole: " + err.Error()
		return
	}
	dir, err := ole.ListDir()
	if err != nil {
		obj.Error = "list ole: " + err.Error()
		return
	}
	for _, file := range dir {
		if file.Name() != "Equation Native" {
			continue
		}
		root := dir[0]
		reader := ole.OpenFile(file, root)
		native, err := io.ReadAll(reader)
		if err != nil {
			obj.Error = "read native: " + err.Error()
			return
		}
		obj.NativeSize = len(native)
		if len(native) < int(oleCbHdr) {
			obj.Error = "native too small"
			return
		}
		obj.HeaderHex = hex.EncodeToString(native[:oleCbHdr])
		cbHdr := binary.LittleEndian.Uint16(native[0:2])
		cbSize := binary.LittleEndian.Uint32(native[8:12])
		if cbHdr == 0 || int(cbHdr) > len(native) {
			obj.Error = fmt.Sprintf("invalid cbHdr=%d", cbHdr)
			return
		}
		end := min(len(native), int(cbHdr)+int(cbSize))
		body := native[cbHdr:end]
		if len(body) >= 3 {
			obj.MtefVersion = int(body[0])
			obj.MtefProduct = int(body[2])
		}
		obj.BodySize = len(body)
		obj.body = append([]byte(nil), body...)
		obj.content = formulaContent(body)
		obj.contentCandidates = formulaContentCandidates(body)
		sum := sha256.Sum256(body)
		obj.BodySHA256 = hex.EncodeToString(sum[:])
		obj.PrefixHex = hex.EncodeToString(body[:min(len(body), 32)])
		obj.Records = recordHistogram(body)
		obj.tail = formulaTail(body)
		obj.tailCandidates = formulaTailCandidates(body)
		obj.TailSize = len(obj.tail)
		if obj.TailSize > 0 {
			tailSum := sha256.Sum256(obj.tail)
			obj.TailSHA256 = hex.EncodeToString(tailSum[:])
			obj.TailRecords = recordHistogram(obj.tail)
		}
		return
	}
	obj.Error = "Equation Native stream not found"
}

func formulaContent(body []byte) []byte {
	for i := mtefRecordStart(body); i < len(body); i++ {
		if i+1 < len(body) && body[i] == 0x01 && (body[i+1] == 0x00 || body[i+1] == 0x01 || body[i+1] == 0x04) && lineHasMeaningfulTail(body[i:]) {
			return body[i:]
		}
		i += recordPayloadLength(body, i)
	}
	return body
}

func formulaTail(body []byte) []byte {
	lastLine := -1
	for i := mtefRecordStart(body); i+1 < len(body); i++ {
		if body[i] == 0x01 && (body[i+1] == 0x00 || body[i+1] == 0x01 || body[i+1] == 0x04) && lineHasMeaningfulTail(body[i:]) {
			lastLine = i
		}
		i += recordPayloadLength(body, i)
	}
	if lastLine < 0 {
		return body
	}
	return body[lastLine:]
}

func formulaContentCandidates(body []byte) [][]byte {
	return formulaCandidates(body, formulaContent(body), false)
}

func formulaTailCandidates(body []byte) [][]byte {
	return formulaCandidates(body, formulaTail(body), false)
}

func formulaCandidates(body []byte, primary []byte, includeLastOnly bool) [][]byte {
	candidates := [][]byte{primary}
	for i := 0; i+2 < len(body); i++ {
		if body[i] == 0x0A && body[i+1] == 0x01 && (body[i+2] == 0x00 || body[i+2] == 0x01 || body[i+2] == 0x04) && lineHasMeaningfulTail(body[i+1:]) {
			candidate := body[i:]
			if includeLastOnly {
				if len(candidates) == 1 {
					candidates = append(candidates, candidate)
				} else {
					candidates[len(candidates)-1] = candidate
				}
			} else {
				candidates = append(candidates, candidate)
			}
		}
		if start := fractionSlotStart(body, i); start >= 0 {
			candidates = append(candidates, body[start:])
		}
		if start := pileInnerLineStart(body, i); start >= 0 {
			candidates = append(candidates, body[start:])
		}
	}
	return uniqueByteSlices(candidates)
}

func pileInnerLineStart(body []byte, i int) int {
	if i+6 >= len(body) || body[i] != 0x04 {
		return -1
	}
	pileLen := recordPayloadLength(body, i)
	start := i + pileLen
	if start < len(body) && body[start] == 0x0A {
		start++
	}
	if start+1 < len(body) && body[start] == 0x01 && lineHasMeaningfulTail(body[start:]) {
		return start
	}
	return -1
}

func fractionSlotStart(body []byte, i int) int {
	if i+4 >= len(body) || body[i] != 0x03 {
		return -1
	}
	options := body[i+1]
	extra := 1
	if options&0x08 != 0 {
		extra += nudgePayloadLength(body, i+1+extra)
	}
	selectorOffset := i + 1 + extra
	if selectorOffset >= len(body) || body[selectorOffset] != 0x0B {
		return -1
	}
	templateLen := recordPayloadLength(body, i)
	start := i + templateLen
	if start+1 < len(body) && body[start] == 0x01 && lineHasMeaningfulTail(body[start:]) {
		return start
	}
	return -1
}

func uniqueByteSlices(items [][]byte) [][]byte {
	out := make([][]byte, 0, len(items))
	seen := map[string]bool{}
	for _, item := range items {
		key := fmt.Sprintf("%p:%d", item, len(item))
		if len(item) == 0 || seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, item)
	}
	if len(out) == 0 {
		return [][]byte{{}}
	}
	return out
}

func mtefRecordStart(body []byte) int {
	if len(body) <= 5 {
		return 0
	}
	if bytes.Equal(body[5:min(len(body), 10)], []byte("DSMT6")) {
		for i := 10; i < len(body); i++ {
			if body[i] == 0 {
				return min(len(body), i+2)
			}
		}
	}
	for i := 5; i < len(body); i++ {
		if body[i] == 0 {
			return i + 1
		}
	}
	return 5
}

func lineHasMeaningfulTail(body []byte) bool {
	for i := 2; i < len(body); i++ {
		tag := body[i]
		if tag == 0x02 || tag == 0x03 || tag == 0x04 || tag == 0x05 {
			return true
		}
		if tag == 0x00 {
			return false
		}
		if tag >= 100 && i+1 < len(body) {
			i += 1 + int(body[i+1])
		}
	}
	return false
}

func writeWorstDumps(outDir, stem, sourcePath, generatedPath string, pairs []PairRow) error {
	source, err := extractObjects(sourcePath)
	if err != nil {
		return err
	}
	generated, err := extractObjects(generatedPath)
	if err != nil {
		return err
	}
	worst := append([]PairRow(nil), pairs...)
	sort.Slice(worst, func(i, j int) bool {
		return worst[i].RecordCosine < worst[j].RecordCosine
	})
	selected := make([]PairRow, 0, min(10, len(worst)))
	seen := map[int]bool{}
	limit := min(10, len(worst))
	for i := 0; i < limit; i++ {
		selected = append(selected, worst[i])
		seen[worst[i].SourceIndex] = true
	}
	for _, sourceIndex := range dumpSourceIndexes() {
		if seen[sourceIndex] {
			continue
		}
		for _, pair := range pairs {
			if pair.SourceIndex == sourceIndex {
				selected = append(selected, pair)
				seen[sourceIndex] = true
				break
			}
		}
	}
	for _, pair := range selected {
		sourceIdx := pair.SourceIndex - 1
		generatedIdx := pair.GeneratedIndex - 1
		if sourceIdx < 0 || sourceIdx >= len(source) || generatedIdx < 0 || generatedIdx >= len(generated) {
			continue
		}
		base := filepath.Join(outDir, fmt.Sprintf("%s_pair_%03d", stem, pair.SourceIndex))
		if err := os.WriteFile(base+"_source_body.hex", []byte(hex.Dump(source[sourceIdx].body)), 0644); err != nil {
			return err
		}
		if err := os.WriteFile(base+"_source_tail.hex", []byte(hex.Dump(source[sourceIdx].tail)), 0644); err != nil {
			return err
		}
		if err := os.WriteFile(base+"_generated_body.hex", []byte(hex.Dump(generated[generatedIdx].body)), 0644); err != nil {
			return err
		}
		if err := os.WriteFile(base+"_generated_tail.hex", []byte(hex.Dump(generated[generatedIdx].tail)), 0644); err != nil {
			return err
		}
		diff := map[string]any{
			"index": pair.Index,
			"sourceIndex": pair.SourceIndex,
			"generatedIndex": pair.GeneratedIndex,
			"latex": pair.Latex,
			"sourceEntry": source[sourceIdx].Entry,
			"generatedEntry": generated[generatedIdx].Entry,
			"sourceRecords": source[sourceIdx].Records,
			"generatedRecords": generated[generatedIdx].Records,
			"sourceTailRecords": source[sourceIdx].TailRecords,
			"generatedTailRecords": generated[generatedIdx].TailRecords,
			"sourcePrefixHex": source[sourceIdx].PrefixHex,
			"generatedPrefixHex": generated[generatedIdx].PrefixHex,
			"recordCosine": pair.RecordCosine,
			"tailRecordCosine": pair.TailRecordCosine,
			"bodySizeRatio": pair.BodySizeRatio,
			"tailSizeRatio": pair.TailSizeRatio,
		}
		data, _ := json.MarshalIndent(diff, "", "  ")
		if err := os.WriteFile(base+"_diff.json", append(data, '\n'), 0644); err != nil {
			return err
		}
	}
	return nil
}

func dumpSourceIndexes() []int {
	raw := strings.TrimSpace(os.Getenv("MTEF_DUMP_SOURCE_INDEXES"))
	if raw == "" {
		return nil
	}
	parts := strings.Split(raw, ",")
	indexes := make([]int, 0, len(parts))
	for _, part := range parts {
		value, err := strconv.Atoi(strings.TrimSpace(part))
		if err == nil && value > 0 {
			indexes = append(indexes, value)
		}
	}
	return indexes
}

type reportFile struct {
	Equations []struct {
		Status string `json:"status"`
		Output string `json:"output"`
		Source string `json:"source"`
	} `json:"equations"`
}

type latexObject struct {
	Latex       string
	ObjectIndex int
}

type requestFile struct {
	Sections []struct {
		Questions []struct {
			Content string `json:"content"`
		} `json:"questions"`
	} `json:"sections"`
}

func reportLatexSequence(path string) ([]latexObject, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var report reportFile
	if err := json.Unmarshal(data, &report); err != nil {
		return nil, err
	}
	out := make([]latexObject, 0, len(report.Equations))
	for _, eq := range report.Equations {
		if eq.Status == "converted" && strings.TrimSpace(eq.Output) != "" {
			out = append(out, latexObject{
				Latex: normalizeLatexKey(eq.Output),
				ObjectIndex: objectIndexFromSource(eq.Source),
			})
		}
	}
	return out, nil
}

func objectIndexFromSource(source string) int {
	match := regexp.MustCompile(`oleObject(\d+)\.bin`).FindStringSubmatch(source)
	if len(match) != 2 {
		return 0
	}
	value, err := strconv.Atoi(match[1])
	if err != nil {
		return 0
	}
	return value
}

func requestLatexSequence(path string) ([]string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var request requestFile
	if err := json.Unmarshal(data, &request); err != nil {
		return nil, err
	}
	text := strings.Builder{}
	for _, section := range request.Sections {
		for _, question := range section.Questions {
			text.WriteString(question.Content)
			text.WriteByte('\n')
		}
	}
	re := regexp.MustCompile(`(?s)\$\$(.+?)\$\$|\$(.+?)\$`)
	matches := re.FindAllStringSubmatch(text.String(), -1)
	out := make([]string, 0, len(matches))
	for _, match := range matches {
		body := match[1]
		if body == "" {
			body = match[2]
		}
		out = append(out, normalizeLatexKey(body))
	}
	return out, nil
}

func normalizeLatexKey(value string) string {
	value = strings.TrimSpace(value)
	value = regexp.MustCompile(`^\\pwmetrics\{[^}]+}\s*`).ReplaceAllString(value, "")
	value = regexp.MustCompile(`^\\pwstyle\{[^}]*}\s*`).ReplaceAllString(value, "")
	value = repairEmptyFractionDenominators(value)
	value = strings.ReplaceAll(value, `\lt `, "<")
	value = strings.ReplaceAll(value, `\gt `, ">")
	value = strings.Join(strings.Fields(value), " ")
	return value
}

func repairEmptyFractionDenominators(value string) string {
	re := regexp.MustCompile(`\\frac\s*\{\s*([^{}]+?)\s*\}\s*\{\s*\}\s*((?:[A-Za-z0-9]+)(?:\s*\\times\s*[A-Za-z0-9]+)*)`)
	return re.ReplaceAllString(value, `\frac { $1 } { $2 }`)
}

func ordinalPairs(sourceLen, generatedLen int) [][2]int {
	n := min(sourceLen, generatedLen)
	pairs := make([][2]int, 0, n)
	for i := 0; i < n; i++ {
		pairs = append(pairs, [2]int{i, i})
	}
	return pairs
}

func latexPairs(sourceLatex []latexObject, generatedLatex []string, sourceLen, generatedLen int) [][2]int {
	byKey := map[string][]int{}
	for i, latex := range generatedLatex {
		byKey[latex] = append(byKey[latex], i)
	}
	used := map[int]bool{}
	pairs := make([][2]int, 0, min(len(sourceLatex), len(generatedLatex)))
	for _, sourceItem := range sourceLatex {
		sourceIndex := sourceItem.ObjectIndex - 1
		if sourceIndex < 0 {
			continue
		}
		queue := byKey[sourceItem.Latex]
		for len(queue) > 0 && used[queue[0]] {
			queue = queue[1:]
		}
		byKey[sourceItem.Latex] = queue
		if len(queue) == 0 {
			continue
		}
		generatedIndex := queue[0]
		byKey[sourceItem.Latex] = queue[1:]
		if sourceIndex < sourceLen && generatedIndex < generatedLen {
			used[generatedIndex] = true
			pairs = append(pairs, [2]int{sourceIndex, generatedIndex})
		}
	}
	return pairs
}

func recordHistogram(body []byte) map[string]int {
	h := map[string]int{}
	for i := 0; i < len(body); i++ {
		tag := body[i]
		h[strconv.Itoa(int(tag))]++
		i += recordPayloadLength(body, i)
	}
	return h
}

func bestCandidateCosine(source, generated [][]byte) float64 {
	best := 0.0
	for _, s := range source {
		for _, g := range generated {
			score := cosine(recordHistogram(s), recordHistogram(g))
			if score > best {
				best = score
			}
		}
	}
	return best
}

func bestCandidateCoreCosine(source, generated [][]byte) float64 {
	best := 0.0
	for _, s := range source {
		for _, g := range generated {
			score := cosine(recordHistogramCore(s), recordHistogramCore(g))
			if score > best {
				best = score
			}
		}
	}
	return best
}

func recordPayloadLength(body []byte, i int) int {
	if i >= len(body) {
		return 0
	}
	tag := body[i]
	switch tag {
	case 0x00, 0x08, 0x0A, 0x0B, 0x0C:
		return 0
	case 0x01:
		if i+1 < len(body) {
			options := body[i+1]
			extra := 1
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+1+extra)
			}
			if options&0x04 != 0 {
				extra++
			}
			if options&0x02 != 0 && i+1+extra < len(body) {
				nStops := int(body[i+1+extra])
				extra += 1 + nStops*3
			}
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x02:
		if i+2 < len(body) {
			options := body[i+1]
			extra := 2
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+extra)
			}
			if options&0x20 == 0 {
				extra += 2
			}
			if options&0x04 != 0 {
				extra++
			}
			if options&0x10 != 0 {
				extra += 2
			}
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x03:
		if i+4 < len(body) {
			options := body[i+1]
			extra := 2
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+extra)
			}
			extra++ // selector
			if i+extra < len(body) {
				variation := body[i+extra]
				extra++
				if variation&0x80 != 0 {
					extra++
				}
			}
			extra++ // template options
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x04:
		if i+3 < len(body) {
			options := body[i+1]
			extra := 2
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+extra)
			}
			extra += 2 // halign + valign
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x05:
		if i+6 < len(body) {
			options := body[i+1]
			extra := 2
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+extra)
			}
			dimOffset := i + extra + 3 // valign + h_just + v_just
			extra += 5                 // valign + h_just + v_just + rows + cols
			if i+extra < len(body) && dimOffset+1 < len(body) {
				rows := int(body[dimOffset])
				cols := int(body[dimOffset+1])
				extra += packedPartitionBytes(rows + 1)
				extra += packedPartitionBytes(cols + 1)
			}
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x06:
		if i+2 < len(body) {
			options := body[i+1]
			extra := 2
			if options&0x08 != 0 {
				extra += nudgePayloadLength(body, i+extra)
			}
			extra++ // embellishment type
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x09:
		if i+3 < len(body) && body[i+2] == 0x50 {
			return 3
		}
		if i+1 < len(body) {
			return 1
		}
	case 0x0D, 0x0E:
		if i+1 < len(body) {
			return 1
		}
	case 0x0F:
		if i+1 < len(body) {
			return 1
		}
	case 0x10:
		if i+1 < len(body) {
			options := body[i+1]
			extra := 1
			if options&0x01 != 0 {
				extra += 8
			} else {
				extra += 6
			}
			if options&0x04 != 0 {
				for i+extra < len(body) && body[i+extra] != 0 {
					extra++
				}
				extra++
			}
			if i+extra < len(body) {
				return extra
			}
		}
	case 0x11:
		end := i + 2
		for end < len(body) && body[end] != 0 {
			end++
		}
		if end < len(body) {
			return end - i
		}
	case 0x13:
		end := i + 1
		for end < len(body) && body[end] != 0 {
			end++
		}
		if end < len(body) {
			return end - i
		}
	default:
		if tag >= 100 && i+1 < len(body) {
			n := int(body[i+1])
			if i+1+n < len(body) {
				return 1 + n
			}
		}
	}
	return 0
}

func nudgePayloadLength(body []byte, i int) int {
	if i+1 >= len(body) {
		return 0
	}
	if body[i] == 0x80 || body[i+1] == 0x80 {
		return 6
	}
	return 2
}

func packedPartitionBytes(size int) int {
	if size <= 0 {
		return 0
	}
	return (size + 3) / 4
}

func recordHistogramCore(body []byte) map[string]int {
	h := map[string]int{}
	for i := 0; i < len(body); i++ {
		tag := body[i]
		if tag == 0xC8 && i+1 < len(body) {
			i++
			continue
		}
		h[strconv.Itoa(int(tag))]++
		i += recordPayloadLength(body, i)
	}
	return h
}

func cosine(a, b map[string]int) float64 {
	if len(a) == 0 || len(b) == 0 {
		return 0
	}
	var dot, aa, bb float64
	for k, av := range a {
		aa += float64(av * av)
		if bv, ok := b[k]; ok {
			dot += float64(av * bv)
		}
	}
	for _, bv := range b {
		bb += float64(bv * bv)
	}
	if aa == 0 || bb == 0 {
		return 0
	}
	return dot / (sqrt(aa) * sqrt(bb))
}

func commonSuffixBytes(a, b []byte) int {
	n := 0
	for n < len(a) && n < len(b) {
		if a[len(a)-1-n] != b[len(b)-1-n] {
			break
		}
		n++
	}
	return n
}

func sqrt(x float64) float64 {
	z := x
	if z == 0 {
		return 0
	}
	for i := 0; i < 16; i++ {
		z = (z + x/z) / 2
	}
	return z
}

func writeCSV(path string, rows []PairRow) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	w := csv.NewWriter(f)
	defer w.Flush()
	_ = w.Write([]string{"index", "sourceIndex", "generatedIndex", "sourceEntry", "generatedEntry", "sourceBodySize", "generatedBodySize", "bodySizeRatio", "bodyHashEqual", "recordCosine", "tailSizeRatio", "tailHashEqual", "tailRecordCosine", "tailCoreRecordCosine", "commonSuffixBytes", "commonSuffixRatio", "alignmentSuspect", "sourceRecordTotal", "generatedRecordTotal", "sourceMtefVersion", "sourceMtefProduct", "latex", "sourceError", "generatedError"})
	for _, r := range rows {
		_ = w.Write([]string{
			strconv.Itoa(r.Index),
			strconv.Itoa(r.SourceIndex),
			strconv.Itoa(r.GeneratedIndex),
			r.SourceEntry,
			r.GeneratedEntry,
			strconv.Itoa(r.SourceBodySize),
			strconv.Itoa(r.GeneratedBodySize),
			fmt.Sprintf("%.6f", r.BodySizeRatio),
			strconv.FormatBool(r.BodyHashEqual),
			fmt.Sprintf("%.6f", r.RecordCosine),
			fmt.Sprintf("%.6f", r.TailSizeRatio),
			strconv.FormatBool(r.TailHashEqual),
			fmt.Sprintf("%.6f", r.TailRecordCosine),
			fmt.Sprintf("%.6f", r.TailCoreRecordCosine),
			strconv.Itoa(r.CommonSuffixBytes),
			fmt.Sprintf("%.6f", r.CommonSuffixRatio),
			strconv.FormatBool(r.AlignmentSuspect),
			strconv.Itoa(r.SourceRecordTotal),
			strconv.Itoa(r.GeneratedRecordTotal),
			strconv.Itoa(r.SourceMtefVersion),
			strconv.Itoa(r.SourceMtefProduct),
			r.Latex,
			r.SourceError,
			r.GeneratedError,
		})
	}
	return nil
}

func readZipEntry(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("entry not found: %s", name)
}

func countReadable(rows []OleObject) int {
	n := 0
	for _, row := range rows {
		if row.BodySize > 0 && row.Error == "" {
			n++
		}
	}
	return n
}

func recordTotal(records map[string]int) int {
	total := 0
	for _, n := range records {
		total += n
	}
	return total
}

func median(values []float64) float64 {
	if len(values) == 0 {
		return 0
	}
	sort.Float64s(values)
	mid := len(values) / 2
	if len(values)%2 == 1 {
		return values[mid]
	}
	return (values[mid-1] + values[mid]) / 2
}

var numberRe = regexp.MustCompile(`\d+`)

func naturalLess(a, b string) bool {
	an := numberRe.FindString(a)
	bn := numberRe.FindString(b)
	if an != "" && bn != "" {
		ai, _ := strconv.Atoi(an)
		bi, _ := strconv.Atoi(bn)
		if ai != bi {
			return ai < bi
		}
	}
	return a < b
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
