package xlsxcache_test

import (
	"archive/zip"
	"bytes"
	"fmt"
	xc "github.com/Alifarid0011/xlsxcache"
	"io"
	"os"
	"strings"
	"testing"
)

const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const relNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const pkgNS = "http://schemas.openxmlformats.org/package/2006/relationships"

func parts(cells string) map[string]string {
	return map[string]string{
		"_rels/.rels":                `<Relationships xmlns="` + pkgNS + `"><Relationship Id="root" Type="` + relNS + `/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml":            `<workbook xmlns="` + ns + `" xmlns:r="` + relNS + `"><sheets><sheet name="Revenue &amp; costs" sheetId="1" state="hidden" r:id="rId7"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="` + pkgNS + `"><Relationship Id="rId7" Type="` + relNS + `/worksheet" Target="worksheets/sheet7.xml"/></Relationships>`,
		"xl/worksheets/sheet7.xml":   `<worksheet xmlns="` + ns + `"><sheetData><row r="1">` + cells + `</row></sheetData></worksheet>`,
	}
}
func archive(t testing.TB, p map[string]string) []byte {
	t.Helper()
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for name, body := range p {
		w, err := z.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = w.Write([]byte(body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}
	return b.Bytes()
}
func inspect(t testing.TB, p map[string]string) (xc.Report, error) {
	t.Helper()
	b := archive(t, p)
	return xc.Inspect(bytes.NewReader(b), int64(len(b)))
}
func TestFormulaCacheKinds(t *testing.T) {
	cases := []struct {
		name, cell                string
		formulas, missing, errors int
	}{
		{"absent", `<c r="B1"><f>1+2</f></c>`, 1, 1, 0},
		{"empty numeric", `<c r="B1"><f>1+2</f><v/></c>`, 1, 1, 0},
		{"zero", `<c r="B1"><f>0</f><v>0</v></c>`, 1, 0, 0},
		{"false", `<c r="B1" t="b"><f>FALSE()</f><v>0</v></c>`, 1, 0, 0},
		{"empty string", `<c r="B1" t="str"><f>IF(1,"","")</f><v/></c>`, 1, 0, 0},
		{"string absent value", `<c r="B1" t="str"><f>""</f></c>`, 1, 1, 0},
		{"error", `<c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c>`, 1, 0, 1},
		{"shared follower", `<c r="B1"><f t="shared" si="0"/></c>`, 1, 1, 0},
		{"array anchor", `<c r="B1"><f t="array" ref="B1:B4">SEQUENCE(4)</f></c>`, 1, 1, 0},
		{"literal ignored", `<c r="B1" t="e"><v>#N/A</v></c>`, 0, 0, 0},
	}
	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			got, err := inspect(t, parts(tt.cell))
			if err != nil {
				t.Fatal(err)
			}
			if got.Sheets != 1 || got.Formulas != tt.formulas || got.Missing != tt.missing || got.Errors != tt.errors {
				t.Fatalf("unexpected report: %+v", got)
			}
			if tt.missing+tt.errors > 0 {
				if len(got.Findings) != 1 || got.Findings[0].Sheet != "Revenue & costs" || got.Findings[0].Cell != "B1" {
					t.Fatalf("lost location: %+v", got)
				}
			}
		})
	}
}
func TestBrokenInputsFail(t *testing.T) {
	cases := []struct {
		name string
		edit func(map[string]string)
	}{
		{"missing root", func(p map[string]string) { delete(p, "_rels/.rels") }},
		{"missing worksheet", func(p map[string]string) { delete(p, "xl/worksheets/sheet7.xml") }},
		{"missing relationship", func(p map[string]string) { p["xl/_rels/workbook.xml.rels"] = `<Relationships xmlns="` + pkgNS + `"/>` }},
		{"external worksheet", func(p map[string]string) {
			p["xl/_rels/workbook.xml.rels"] = strings.Replace(p["xl/_rels/workbook.xml.rels"], `Target=`, `TargetMode="External" Target=`, 1)
		}},
		{"trailing garbage", func(p map[string]string) { p["xl/worksheets/sheet7.xml"] += "<broken" }},
		{"wrong namespace", func(p map[string]string) {
			p["xl/worksheets/sheet7.xml"] = strings.ReplaceAll(p["xl/worksheets/sheet7.xml"], ns, "urn:wrong")
		}},
		{"missing cell reference", func(p map[string]string) {
			p["xl/worksheets/sheet7.xml"] = strings.ReplaceAll(p["xl/worksheets/sheet7.xml"], `r="B1"`, "")
		}},
		{"path traversal", func(p map[string]string) {
			p["xl/_rels/workbook.xml.rels"] = strings.ReplaceAll(p["xl/_rels/workbook.xml.rels"], "worksheets/sheet7.xml", "../../outside.xml")
		}},
	}
	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			p := parts(`<c r="B1"><f>1+2</f></c>`)
			tt.edit(p)
			if _, err := inspect(t, p); err == nil {
				t.Fatal("expected invalid-input error")
			}
		})
	}
	if _, err := xc.Inspect(bytes.NewReader([]byte("not zip")), 7); err == nil {
		t.Fatal("accepted non-ZIP")
	}
}
func TestStrictAndRelocatedWorkbook(t *testing.T) {
	p := parts(`<c r="B1"><f>1+2</f></c>`)
	q := map[string]string{}
	for k, v := range p {
		k = strings.ReplaceAll(k, "xl/", "custom/")
		v = strings.ReplaceAll(v, "xl/", "custom/")
		v = strings.ReplaceAll(v, ns, "http://purl.oclc.org/ooxml/spreadsheetml/main")
		v = strings.ReplaceAll(v, relNS, "http://purl.oclc.org/ooxml/officeDocument/relationships")
		q[k] = v
	}
	got, err := inspect(t, q)
	if err != nil {
		t.Fatal(err)
	}
	if got.Missing != 1 {
		t.Fatalf("lost formula: %+v", got)
	}
}
func TestFindingLimitKeepsCounts(t *testing.T) {
	var cells strings.Builder
	for i := 1; i <= 1005; i++ {
		fmt.Fprintf(&cells, `<c r="B%d"><f>1</f></c>`, i)
	}
	got, err := inspect(t, parts(cells.String()))
	if err != nil {
		t.Fatal(err)
	}
	if got.Missing != 1005 || len(got.Findings) != 1000 || !got.Truncated {
		t.Fatalf("unexpected bounded report: %+v", got)
	}
}
func FuzzInspect(f *testing.F) {
	f.Add([]byte("not zip"))
	f.Add(archive(f, parts(`<c r="B1"><f>1</f></c>`)))
	f.Fuzz(func(t *testing.T, b []byte) { _, _ = xc.Inspect(bytes.NewReader(b), int64(len(b))) })
}

func TestArchiveBoundsAndAmbiguity(t *testing.T) {
	for _, tt := range []struct {
		name   string
		header zip.FileHeader
		want   string
	}{
		{"XML size", zip.FileHeader{Name: "_rels/.rels", UncompressedSize64: 65 << 20}, "XML bytes"},
		{"total expansion", zip.FileHeader{Name: "unused.bin", UncompressedSize64: 513 << 20}, "expansion"},
		{"encrypted", zip.FileHeader{Name: "unused.bin", Flags: 1}, "encrypted"},
		{"unsafe name", zip.FileHeader{Name: "../unused.bin"}, "part name"},
	} {
		t.Run(tt.name, func(t *testing.T) {
			var b bytes.Buffer
			z := zip.NewWriter(&b)
			if _, err := z.CreateRaw(&tt.header); err != nil {
				t.Fatal(err)
			}
			if err := z.Close(); err != nil {
				t.Fatal(err)
			}
			_, err := xc.Inspect(bytes.NewReader(b.Bytes()), int64(b.Len()))
			if err == nil || !strings.Contains(err.Error(), tt.want) {
				t.Fatalf("got %v; want %s", err, tt.want)
			}
		})
	}
	for _, n := range []int64{-1, 0, 257 << 20} {
		if _, err := xc.Inspect(bytes.NewReader(nil), n); err == nil {
			t.Fatalf("accepted size %d", n)
		}
	}
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for i := 0; i < 2; i++ {
		if _, err := z.Create("_rels/.rels"); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := xc.Inspect(bytes.NewReader(b.Bytes()), int64(b.Len())); err == nil || !strings.Contains(err.Error(), "duplicate") {
		t.Fatalf("duplicate accepted: %v", err)
	}
}
func TestMalformedAfterFindingsReturnsNoReport(t *testing.T) {
	p := parts(`<c r="B1"><f>1</f></c>`)
	p["xl/worksheets/sheet7.xml"] += "<extra/>"
	got, err := inspect(t, p)
	if err == nil || got.Formulas != 0 || got.Findings != nil {
		t.Fatalf("partial report leaked: %+v, %v", got, err)
	}
}
func TestMissingSheetDataFails(t *testing.T) {
	p := parts("")
	p["xl/worksheets/sheet7.xml"] = `<worksheet xmlns="` + ns + `"/>`
	if _, err := inspect(t, p); err == nil {
		t.Fatal("worksheet without sheetData accepted")
	}
}

func TestArchiveChecksumAndEntryCount(t *testing.T) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	body := parts("")["_rels/.rels"]
	w, err := z.CreateHeader(&zip.FileHeader{Name: "_rels/.rels", Method: zip.Store})
	if err != nil {
		t.Fatal(err)
	}
	if _, err = io.WriteString(w, body); err != nil {
		t.Fatal(err)
	}
	if err = z.Close(); err != nil {
		t.Fatal(err)
	}
	data := b.Bytes()
	i := bytes.Index(data, []byte("Relationships"))
	if i < 0 {
		t.Fatal("missing stored XML")
	}
	data[i] = 'X'
	if _, err := xc.Inspect(bytes.NewReader(data), int64(len(data))); err == nil {
		t.Fatal("accepted corrupt ZIP checksum")
	}
	b.Reset()
	z = zip.NewWriter(&b)
	for i := 0; i < 10001; i++ {
		if _, err := z.Create(fmt.Sprintf("p%d", i)); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := xc.Inspect(bytes.NewReader(b.Bytes()), int64(b.Len())); err == nil || !strings.Contains(err.Error(), "entries") {
		t.Fatalf("entry limit: %v", err)
	}
}

func TestMalformedWorksheetCannotLookClean(t *testing.T) {
	for _, cells := range []string{
		`<c r="B1"><f>1</f><x:v xmlns:x="urn:foreign">1</x:v></c>`,
		`<c r="B1"><x:f xmlns:x="urn:foreign">1</x:f><v>1</v></c>`,
		`<c r="B1"><f>1</f><v>1</v><v>2</v></c>`,
	} {
		if _, err := inspect(t, parts(cells)); err == nil {
			t.Fatalf("accepted ambiguous formula cell: %s", cells)
		}
	}
	p := parts("")
	p["xl/worksheets/sheet7.xml"] = `<worksheet xmlns="` + ns + `"><sheetData><c r="A1"><f>1</f></c></sheetData></worksheet>`
	if _, err := inspect(t, p); err == nil {
		t.Fatal("misplaced formula cell silently skipped")
	}
}

func TestExcelizeGeneratedWorkbook(t *testing.T) {
	f, err := os.Open("testdata/missing-cache.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	got, err := xc.Inspect(f, info.Size())
	if err != nil {
		t.Fatal(err)
	}
	if got.Sheets != 1 || got.Formulas != 2 || got.Missing != 2 || got.Errors != 0 || len(got.Findings) != 2 || got.Findings[0].Cell != "C2" || got.Findings[1].Cell != "D2" {
		t.Fatalf("unexpected real-writer report: %+v", got)
	}
}

func TestLibreOfficeRecalculatedWorkbook(t *testing.T) {
	f, err := os.Open("testdata/cached.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	got, err := xc.Inspect(f, info.Size())
	if err != nil {
		t.Fatal(err)
	}
	if got.Sheets != 1 || got.Formulas != 2 || got.Missing != 0 || got.Errors != 0 || len(got.Findings) != 0 {
		t.Fatalf("unexpected recalculated report: %+v", got)
	}
}

func TestExtensionAttributesCannotOverrideStandardAttributes(t *testing.T) {
	p := parts(`<c r="B1" t="e" x:t="n" x:r="C1" xmlns:x="urn:extension"><f>1/0</f><v>#DIV/0!</v></c>`)
	got, err := inspect(t, p)
	if err != nil {
		t.Fatal(err)
	}
	if got.Errors != 1 || len(got.Findings) != 1 || got.Findings[0].Cell != "B1" {
		t.Fatalf("extension overrode cell attributes: %+v", got)
	}
	p = parts(`<c r="B1"><f>1</f></c>`)
	p["xl/_rels/workbook.xml.rels"] = strings.Replace(p["xl/_rels/workbook.xml.rels"], `Target="worksheets/sheet7.xml"`, `Target="worksheets/sheet7.xml" x:Target="worksheets/clean.xml" xmlns:x="urn:extension"`, 1)
	p["xl/worksheets/clean.xml"] = `<worksheet xmlns="` + ns + `"><sheetData/></worksheet>`
	got, err = inspect(t, p)
	if err != nil {
		t.Fatal(err)
	}
	if got.Missing != 1 {
		t.Fatalf("extension redirected worksheet: %+v", got)
	}
}
