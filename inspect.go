// Package xlsxcache checks saved formula results without evaluating formulas.
package xlsxcache

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"path"
	"strings"
)

const (
	maxArchive  = 256 << 20
	maxExpanded = 512 << 20
	maxXML      = 64 << 20
	maxEntries  = 10000
	maxFindings = 1000
	mainNS      = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	strictNS    = "http://purl.oclc.org/ooxml/spreadsheetml/main"
	relNS       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	strictRelNS = "http://purl.oclc.org/ooxml/officeDocument/relationships"
	packageNS   = "http://schemas.openxmlformats.org/package/2006/relationships"
)

// Finding identifies a formula with a missing cache or a cached error.
// Formula text and cached values are deliberately omitted.
type Finding struct {
	Sheet string `json:"sheet"`
	Cell  string `json:"cell"`
	Code  string `json:"code"`
}

// Report counts every inspected explicit formula cell, including shared followers.
// Findings contains at most the first 1000 locations in workbook/XML order.
// A report with no findings does not establish freshness or correctness.
type Report struct {
	Sheets    int       `json:"sheets"`
	Formulas  int       `json:"formulas"`
	Missing   int       `json:"missing"`
	Errors    int       `json:"errors"`
	Findings  []Finding `json:"findings"`
	Truncated bool      `json:"truncated"`
}

type workbook struct {
	Sheets []struct {
		Name  string     `xml:"-"`
		Attrs []xml.Attr `xml:",any,attr"`
	} `xml:"sheets>sheet"`
}
type relationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
	Mode   string `xml:"TargetMode,attr"`
}
type archive map[string]*zip.File

// Inspect reads an unencrypted XLSX/XLSM package. It never writes, executes
// formulas or macros, or fetches links. On invalid input it returns a zero Report
// and an error; callers must not treat a partially inspected workbook as clean.
// Input is limited to 256 MiB, 10000 ZIP entries, 512 MiB declared expansion,
// and 64 MiB per inspected XML part. Non-worksheet sheets are skipped.
func Inspect(r io.ReaderAt, size int64) (Report, error) {
	report, err := inspect(r, size)
	if err != nil {
		return Report{}, err
	}
	return report, nil
}
func inspect(r io.ReaderAt, size int64) (Report, error) {
	report := Report{Findings: make([]Finding, 0)}
	if size <= 0 || size > maxArchive {
		return report, fmt.Errorf("archive size must be between 1 and %d bytes", maxArchive)
	}
	z, err := zip.NewReader(r, size)
	if err != nil {
		return report, fmt.Errorf("open ZIP: %w", err)
	}
	if len(z.File) > maxEntries {
		return report, fmt.Errorf("ZIP exceeds %d entries", maxEntries)
	}
	a := make(archive, len(z.File))
	var expanded uint64
	for _, f := range z.File {
		if strings.HasSuffix(f.Name, "/") {
			continue
		}
		if f.Name == "" || strings.Contains(f.Name, `\`) || strings.HasPrefix(f.Name, "/") || path.Clean(f.Name) != f.Name || strings.HasPrefix(f.Name, "../") {
			return report, fmt.Errorf("invalid ZIP part name %q", f.Name)
		}
		if _, ok := a[f.Name]; ok {
			return report, fmt.Errorf("duplicate ZIP part %q", f.Name)
		}
		if f.Flags&1 != 0 {
			return report, fmt.Errorf("encrypted ZIP part %q is unsupported", f.Name)
		}
		if f.UncompressedSize64 > maxExpanded-expanded {
			return report, fmt.Errorf("ZIP expansion exceeds %d bytes", maxExpanded)
		}
		expanded += f.UncompressedSize64
		a[f.Name] = f
	}
	root, err := a.relationships("")
	if err != nil {
		return report, err
	}
	var bookPath string
	for _, r := range root {
		if relationshipType(r.Type, "officeDocument") {
			if bookPath != "" {
				return report, fmt.Errorf("multiple workbook relationships")
			}
			bookPath, err = resolve("", r)
			if err != nil {
				return report, err
			}
		}
	}
	if bookPath == "" {
		return report, fmt.Errorf("missing workbook relationship")
	}
	var book workbook
	err = a.xml(bookPath, "workbook", func(d *xml.Decoder, start xml.StartElement) error { return d.DecodeElement(&book, &start) })
	if err != nil {
		return report, err
	}
	if len(book.Sheets) == 0 {
		return report, fmt.Errorf("workbook has no sheets")
	}
	rels, err := a.relationships(bookPath)
	if err != nil {
		return report, err
	}
	for _, s := range book.Sheets {
		attrs, err := plainAttributes(s.Attrs)
		if err != nil {
			return report, err
		}
		s.Name = attrs["name"]
		var id string
		for _, attr := range s.Attrs {
			if attr.Name.Local == "id" && (attr.Name.Space == relNS || attr.Name.Space == strictRelNS) {
				id = attr.Value
			}
		}
		r, ok := rels[id]
		if !ok || id == "" || s.Name == "" {
			return report, fmt.Errorf("sheet %q has no valid relationship", s.Name)
		}
		// Chart sheets and other standard non-grid sheets have no formula cells.
		if !relationshipType(r.Type, "worksheet") {
			if relationshipType(r.Type, "chartsheet") || relationshipType(r.Type, "dialogsheet") || relationshipType(r.Type, "macrosheet") {
				continue
			}
			return report, fmt.Errorf("sheet %q has unsupported relationship type %q", s.Name, r.Type)
		}
		target, err := resolve(bookPath, r)
		if err != nil {
			return report, err
		}
		err = a.xml(target, "worksheet", func(d *xml.Decoder, start xml.StartElement) error { return scanSheet(d, start, s.Name, &report) })
		if err != nil {
			return report, err
		}
		report.Sheets++
	}
	if report.Sheets == 0 {
		return report, fmt.Errorf("workbook has no worksheets")
	}
	return report, nil
}
func relationshipType(t, kind string) bool { return t == relNS+"/"+kind || t == strictRelNS+"/"+kind }
func resolve(source string, r relationship) (string, error) {
	if r.Mode != "" && r.Mode != "Internal" {
		return "", fmt.Errorf("external relationship %q is unsupported", r.ID)
	}
	u, err := url.Parse(r.Target)
	if err != nil || u.IsAbs() || u.Host != "" || u.RawQuery != "" || u.Fragment != "" || u.Path == "" || strings.Contains(u.Path, `\`) {
		return "", fmt.Errorf("invalid relationship target %q", r.Target)
	}
	var target string
	if strings.HasPrefix(u.Path, "/") {
		target = path.Clean(strings.TrimPrefix(u.Path, "/"))
	} else {
		target = path.Join(path.Dir(source), u.Path)
	}
	if target == "." || target == ".." || strings.HasPrefix(target, "../") || strings.HasPrefix(target, "/") {
		return "", fmt.Errorf("relationship target escapes package: %q", r.Target)
	}
	return target, nil
}
func (a archive) relationships(source string) (map[string]relationship, error) {
	name := "_rels/.rels"
	if source != "" {
		name = path.Join(path.Dir(source), "_rels", path.Base(source)+".rels")
	}
	var doc struct {
		Items []relationship `xml:"Relationship"`
	}
	err := a.xml(name, "Relationships", func(d *xml.Decoder, start xml.StartElement) error { return d.DecodeElement(&doc, &start) })
	if err != nil {
		return nil, err
	}
	result := make(map[string]relationship, len(doc.Items))
	for _, r := range doc.Items {
		if _, ok := result[r.ID]; ok || r.ID == "" {
			return nil, fmt.Errorf("%s: duplicate or empty relationship ID %q", name, r.ID)
		}
		result[r.ID] = r
	}
	return result, nil
}

// xml consumes a single complete XML document and reads through EOF, including
// the ZIP checksum. Limiting actual bytes also bounds forged size declarations.
func (a archive) xml(name, root string, read func(*xml.Decoder, xml.StartElement) error) error {
	f, ok := a[name]
	if !ok {
		return fmt.Errorf("missing ZIP part %q", name)
	}
	if f.UncompressedSize64 > maxXML {
		return fmt.Errorf("%s exceeds %d XML bytes", name, maxXML)
	}
	r, err := f.Open()
	if err != nil {
		return fmt.Errorf("%s: %w", name, err)
	}
	defer r.Close()
	limited := &io.LimitedReader{R: r, N: maxXML + 1}
	d := xml.NewDecoder(limited)
	err = readDocument(d, root, read)
	if limited.N <= 0 {
		return fmt.Errorf("%s exceeds %d XML bytes", name, maxXML)
	}
	if err != nil {
		return fmt.Errorf("%s: %w", name, err)
	}
	return nil
}
func readDocument(d *xml.Decoder, root string, read func(*xml.Decoder, xml.StartElement) error) error {
	seen := false
	for {
		tok, err := d.Token()
		if err == io.EOF {
			if !seen {
				return fmt.Errorf("missing XML root %s", root)
			}
			return nil
		}
		if err != nil {
			return err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			namespaceOK := v.Name.Space == mainNS || v.Name.Space == strictNS
			if root == "Relationships" {
				namespaceOK = v.Name.Space == packageNS
			}
			if seen || v.Name.Local != root || !namespaceOK {
				return fmt.Errorf("unexpected XML root %s in namespace %q", v.Name.Local, v.Name.Space)
			}
			seen = true
			if err := read(d, v); err != nil {
				return err
			}
		case xml.CharData:
			if strings.TrimSpace(string(v)) != "" {
				return fmt.Errorf("text outside XML root")
			}
		}
	}
}
func scanSheet(d *xml.Decoder, root xml.StartElement, sheet string, report *Report) error {
	stack := []xml.Name{root.Name}
	seenData := false
	for len(stack) > 0 {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			if len(stack) == 1 && v.Name.Local == "sheetData" {
				if seenData {
					return fmt.Errorf("duplicate sheetData")
				}
				seenData = true
			}
			if len(stack) == 3 && stack[1].Local == "sheetData" && stack[2].Local == "row" && v.Name.Local == "c" {
				if v.Name.Space != root.Name.Space {
					return fmt.Errorf("cell uses an unsupported namespace")
				}
				c, err := readCell(d, v, root.Name.Space)
				if err != nil {
					return err
				}
				if !c.Formula {
					continue
				}
				if !cellReference(c.Ref) {
					return fmt.Errorf("formula cell has invalid reference %q", c.Ref)
				}
				report.Formulas++
				code := ""
				if c.Value == nil || (c.Type != "str" && strings.TrimSpace(*c.Value) == "") {
					report.Missing++
					code = "missing_cache"
				} else if c.Type == "e" {
					report.Errors++
					code = "cached_error"
				}
				if code != "" {
					if len(report.Findings) < maxFindings {
						report.Findings = append(report.Findings, Finding{Sheet: sheet, Cell: c.Ref, Code: code})
					} else {
						report.Truncated = true
					}
				}
			} else {
				if v.Name.Space == root.Name.Space && (v.Name.Local == "c" || v.Name.Local == "f" || (v.Name.Local == "row" && (len(stack) != 2 || stack[1].Local != "sheetData")) || (v.Name.Local == "sheetData" && len(stack) != 1)) {
					return fmt.Errorf("misplaced worksheet element %s", v.Name.Local)
				}
				if (v.Name.Local == "sheetData" || v.Name.Local == "row") && v.Name.Space != root.Name.Space {
					return fmt.Errorf("worksheet structure uses an unsupported namespace")
				}
				stack = append(stack, v.Name)
			}
		case xml.EndElement:
			stack = stack[:len(stack)-1]
		}
	}
	if !seenData {
		return fmt.Errorf("worksheet has no sheetData")
	}
	return nil
}
func cellReference(s string) bool {
	i, col := 0, 0
	for i < len(s) && s[i] >= 'A' && s[i] <= 'Z' {
		col = col*26 + int(s[i]-'A'+1)
		if col > 16384 {
			return false
		}
		i++
	}
	if col == 0 || i == len(s) || s[i] == '0' {
		return false
	}
	row := 0
	for ; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
		row = row*10 + int(s[i]-'0')
		if row > 1048576 {
			return false
		}
	}
	return row > 0
}

// readCell preserves element presence and rejects ambiguous result elements.
// encoding/xml field tags alone match local names across namespaces.
type cell struct {
	Ref     string
	Type    string
	Formula bool
	Value   *string
}

func readCell(d *xml.Decoder, start xml.StartElement, namespace string) (cell, error) {
	var raw struct {
		Children []struct {
			XMLName xml.Name
			Text    string `xml:",chardata"`
		} `xml:",any"`
	}
	if err := d.DecodeElement(&raw, &start); err != nil {
		return cell{}, err
	}
	attrs, err := plainAttributes(start.Attr)
	if err != nil {
		return cell{}, err
	}
	c := cell{Ref: attrs["r"], Type: attrs["t"]}
	for _, child := range raw.Children {
		kind := child.XMLName.Local
		if kind != "f" && kind != "v" {
			continue
		}
		if child.XMLName.Space != namespace {
			return c, fmt.Errorf("cell %q: unsupported namespace for %s", c.Ref, kind)
		}
		if kind == "f" {
			if c.Formula {
				return c, fmt.Errorf("cell %q: duplicate formula", c.Ref)
			}
			c.Formula = true
		} else {
			if c.Value != nil {
				return c, fmt.Errorf("cell %q: duplicate cached value", c.Ref)
			}
			value := child.Text
			c.Value = &value
		}
	}
	return c, nil
}

// OOXML's core attributes are unqualified. Extension attributes with the same
// local name must not overwrite them (encoding/xml struct tags allow that).
func plainAttributes(attrs []xml.Attr) (map[string]string, error) {
	result := make(map[string]string)
	for _, a := range attrs {
		if a.Name.Space != "" {
			continue
		}
		if _, exists := result[a.Name.Local]; exists {
			return nil, fmt.Errorf("duplicate attribute %q", a.Name.Local)
		}
		result[a.Name.Local] = a.Value
	}
	return result, nil
}
func (r *relationship) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	if start.Name.Space != packageNS {
		return fmt.Errorf("unsupported relationship namespace %q", start.Name.Space)
	}
	attrs, err := plainAttributes(start.Attr)
	if err != nil {
		return err
	}
	*r = relationship{ID: attrs["Id"], Type: attrs["Type"], Target: attrs["Target"], Mode: attrs["TargetMode"]}
	return d.Skip()
}
