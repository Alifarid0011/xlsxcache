package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func fixture(t *testing.T, cell string) string {
	t.Helper()
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for name, body := range map[string]string{
		"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="a" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml":            `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sales" r:id="b"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="b" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		"xl/worksheets/sheet1.xml":   `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row>` + cell + `</row></sheetData></worksheet>`,
	} {
		w, err := z.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := io.WriteString(w, body); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}
	p := filepath.Join(t.TempDir(), "book.xlsx")
	if err := os.WriteFile(p, b.Bytes(), 0600); err != nil {
		t.Fatal(err)
	}
	return p
}
func TestRunExitCodesAndJSON(t *testing.T) {
	cases := []struct {
		name, cell string
		want       int
	}{
		{"missing", `<c r="B1"><f>SUM(A1:A3)</f></c>`, 1},
		{"cached error", `<c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c>`, 1},
		{"cached zero", `<c r="B1"><f>0</f><v>0</v></c>`, 0},
	}
	for _, tt := range cases {
		t.Run(tt.name, func(t *testing.T) {
			p := fixture(t, tt.cell)
			var out, errout bytes.Buffer
			got := run([]string{"--json", p}, &out, &errout)
			if got != tt.want || errout.Len() != 0 {
				t.Fatalf("exit=%d stderr=%s", got, errout.String())
			}
			var doc struct {
				SchemaVersion int    `json:"schema_version"`
				File          string `json:"file"`
				Freshness     string `json:"freshness"`
				Formulas      int    `json:"formulas"`
			}
			if err := json.Unmarshal(out.Bytes(), &doc); err != nil {
				t.Fatal(err)
			}
			if doc.SchemaVersion != 1 || doc.File != p || doc.Freshness != "not_checked" || doc.Formulas != 1 {
				t.Fatalf("bad JSON: %s", out.String())
			}
		})
	}
}
func TestRunErrorsAndHelp(t *testing.T) {
	bad := filepath.Join(t.TempDir(), "broken.xlsx")
	if err := os.WriteFile(bad, []byte("bad zip"), 0600); err != nil {
		t.Fatal(err)
	}
	for _, args := range [][]string{nil, {"--unknown"}, {"a", "b"}, {"--json", "missing.xlsx"}, {"--json", bad}} {
		var out, errout bytes.Buffer
		if got := run(args, &out, &errout); got != 2 || out.Len() != 0 || errout.Len() == 0 {
			t.Fatalf("args=%v exit=%d out=%q err=%q", args, got, out.String(), errout.String())
		}
	}
	for _, arg := range []string{"--help", "--version"} {
		var out, errout bytes.Buffer
		if got := run([]string{arg}, &out, &errout); got != 0 || out.Len() == 0 {
			t.Fatalf("%s exit=%d out=%q", arg, got, out.String())
		}
	}
	p := fixture(t, `<c r="B1"><f>1</f></c>`)
	var out, errout bytes.Buffer
	if got := run([]string{p}, &out, &errout); got != 1 || !strings.Contains(out.String(), `"Sales"!B1`) || !strings.Contains(out.String(), "freshness not checked") {
		t.Fatalf("exit=%d out=%s err=%s", got, &out, &errout)
	}
}

type brokenWriter struct{}

func (brokenWriter) Write([]byte) (int, error) { return 0, errors.New("output unavailable") }
func TestRunOutputFailure(t *testing.T) {
	p := fixture(t, `<c r="A1"><f>1</f><v>1</v></c>`)
	for _, args := range [][]string{{p}, {"--json", p}, {"--version"}, {"--help"}} {
		var errout bytes.Buffer
		if got := run(args, brokenWriter{}, &errout); got != 2 {
			t.Fatalf("args=%v exit=%d", args, got)
		}
	}
}
