# xlsxcache

**Catch missing Excel formula results before your data pipeline reads them.**

`xlsxcache` checks saved formula caches in XLSX/XLSM files, reports worksheet and
cell locations, and returns an exit code you can use in CI. It never edits the
workbook, evaluates formulas, launches Excel, or makes network requests.
The CLI and Go package use only the Go standard library.

```text
$ xlsxcache testdata/missing-cache.xlsx
"testdata/missing-cache.xlsx": 1 worksheets, 2 formula cells, 2 missing caches, 0 cached errors
  "Sales"!C2: missing_cache
  "Sales"!D2: missing_cache
Cache freshness not checked; saved results may be stale.
```

Exit status: `1`. Both cells contain formulas but have no saved results.

## Install

With Go 1.25 or newer:

```sh
go install github.com/Alifarid0011/xlsxcache/cmd/xlsxcache@v0.1.0
xlsxcache --help
```

Or download a binary from [Releases](https://github.com/Alifarid0011/xlsxcache/releases).
Add your Go binary directory (`go env GOPATH` plus `/bin`, unless you set `GOBIN`)
to `PATH` if needed.

## Why check caches?

A workbook can contain `SUM(...)` without containing its calculated result.
Readers that consume saved values can then see blanks or old values.
[Microsoft documents the `<v>` element as the last calculated formula result](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas).
[OpenPyXL's `data_only` reads saved results](https://openpyxl.readthedocs.io/en/3.1/tutorial.html),
and [Excelize distinguishes writing formulas from calculating them](https://xuri.me/excelize/en/cell.html#set-cell-formula).

This check is for developers receiving workbook exports or generating workbooks
before an unattended import: for example, Excelize → XLSX → Sheetframe, pandas,
or another saved-value reader. It catches **absent results and cached formula
errors**. A successful check does **not** prove that existing results are fresh,
that the formulas are correct, or that a particular importer supports the file.

## Use in scripts

```sh
# The import only runs if the check exits 0.
xlsxcache incoming.xlsx && ./import-sales incoming.xlsx

# Flags precede the path. Shell redirection only writes the report.
xlsxcache --json incoming.xlsx > cache-report.json

# Paths beginning with a dash:
xlsxcache --json -- ./-incoming.xlsx
```

| Exit | Meaning |
| --- | --- |
| `0` | No findings among inspected formula cells; freshness not checked |
| `1` | Missing caches or cached formula errors found |
| `2` | Bad arguments, unreadable/unsupported/malformed input, or output failure |

Reports go to stdout; operational errors go to stderr. Invalid input produces no
report. A report is assembled before writing; a failed output stream can still
contain a partial write. Do not consume a report after exit `2`.
One workbook path is required; stdin, glob expansion inside the tool, and batch
aggregation are outside this version's scope.

JSON uses `schema_version: 1`. `file` is the supplied path. `freshness` is always
`"not_checked"`. Counts cover all inspected cells, even when locations are truncated:

```json
{
  "schema_version": 1,
  "file": "testdata/missing-cache.xlsx",
  "freshness": "not_checked",
  "sheets": 1,
  "formulas": 2,
  "missing": 2,
  "errors": 0,
  "findings": [
    {"sheet": "Sales", "cell": "C2", "code": "missing_cache"},
    {"sheet": "Sales", "cell": "D2", "code": "missing_cache"}
  ],
  "truncated": false
}
```

Formula text and cached values are omitted. File paths and worksheet names remain
in reports. Finding order follows workbook sheet order, then cell XML order.

## A reproducible example

[testdata/missing-cache.xlsx](testdata/missing-cache.xlsx) was generated with
Excelize v2.11.0. It contains a net sale of 100, `C2 = B2*0.2` for tax, and
`D2 = B2+C2` for total. The formulas have no calculated caches.
The [generator](examples/excelize/main.go) is a separate module so Excelize is
not a dependency of the checker:

```sh
cd examples/excelize
go run .
xlsxcache missing-cache.xlsx
```

The generator creates a synthetic workbook in its current directory. It contains
no personal or production data. All checker tests run without Excelize, Python,
Office, or network access.

The same workbook recalculated by LibreOffice 26.2.6.3 is included as
[testdata/cached.xlsx](testdata/cached.xlsx). It has saved values of 20 and 120
and passes with zero findings; both cases are regression-tested.

To populate caches, recalculate with a compatible spreadsheet engine and save a
**copy**, then check the copy. Excel, LibreOffice, Excelize's supported calculation
APIs, and other engines have different compatibility boundaries. `xlsxcache`
does not select an engine or repair missing results.

## Go API

```go
f, err := os.Open("incoming.xlsx")
if err != nil { return err }
defer f.Close()
info, err := f.Stat()
if err != nil { return err }
report, err := xlsxcache.Inspect(f, info.Size())
if err != nil { return err }
if report.Missing+report.Errors > 0 {
    return fmt.Errorf("formula cache preflight failed: %d missing, %d errors",
        report.Missing, report.Errors)
}
```

Import `github.com/Alifarid0011/xlsxcache`. The reader must be non-nil and support
`io.ReaderAt`; inspection does not close it. On error, the report is zero-valued.

## Exact scope and limits

- Checks explicit worksheet cells with `<f>` in Transitional and Strict OOXML,
  including hidden/very-hidden worksheets and empty shared-formula follower
  elements. Resolves workbook and worksheet relationship targets rather than
  guessing `sheet1.xml`. Standard chart/dialog/macro sheets are not inspected.
- `missing_cache`: no `<v>`, or an empty/whitespace-only non-string `<v>`.
  An explicit empty `<v/>` with `t="str"` is a legitimate empty-string result.
  Numeric zero and boolean false are present results.
- `cached_error`: a formula has a nonempty cached value and cell type `t="e"`.
  Literal error cells without `<f>` are outside this check.
- Only cells carrying `<f>` are counted. Array anchors are inspected; spill/array
  dependent cells without `<f>`, formulas in names/tables, external data caches,
  calculation chains, macro formulas and calculation settings are not audited.
- Formula cells require explicit valid A1 addresses. Although OOXML permits
  omitted cell addresses, this version rejects formula cells without them so
  findings cannot carry guessed locations.
- No numeric/type validation of nonempty caches, shared-string lookup, formula
  parsing, full OOXML schema validation, dependency analysis, staleness detection,
  or malware verdict. Macros and external links are never executed or fetched.
- Legacy XLS/XLSB, encrypted workbooks and non-workbook ZIPs are unsupported.
  Malformed inspected XML, missing worksheet parts and ambiguous ZIP entries
  return an error. Unrelated archive parts are not fully validated.
- Maximum input: 256 MiB; 10,000 ZIP entries; 512 MiB declared total expanded
  file sizes; 64 MiB actual bytes per inspected XML part. XML is parsed one cell
  at a time; one large XML token/cell can still allocate up to that limit.
  Only the first 1,000 finding locations are retained. Counts remain complete
  and `truncated` becomes `true`.

## Alternatives and the tradeoff

| Need | Existing option | Where xlsxcache fits |
| --- | --- | --- |
| Filter, join, aggregate, convert tabular data | [Miller](https://github.com/johnkerl/miller), [DuckDB Excel](https://duckdb.org/docs/current/core_extensions/excel), [qsv](https://github.com/dathere/qsv) | This is a formula-cache preflight, not another query/conversion language |
| Compare exports by row key | [csv-diff](https://github.com/simonw/csv-diff) | A different problem; xlsxcache does not diff tables |
| Recalculate caches, including a read-only freshness check | [Formualizer](https://www.formualizer.dev/docs/recalc-cli) | Use it for supported formula evaluation; xlsxcache only examines saved structure |
| Interactive formula/dependency audit | [BaseToolbox Formula Auditor](https://basetoolbox.com/excel-formula-auditor/) | Broader browser workflow; xlsxcache provides a small offline CLI and Go API |
| Workbook structure report through an API | [XLSX Inspector](https://xlsx.lifestep.io/) | Broader structural report; xlsxcache stays local and focuses on formula-cache locations |

These tools already solve substantial parts of spreadsheet workflows. The reason
for this project is a narrow, dependency-free check for scripts and Go programs,
not a claim that formula auditing is new or that other languages cannot do it.
If your pipeline already has an equivalent check, you do not need this tool.

## Development

```sh
go test ./...
go test -race -cover ./...    # requires a C compiler
go vet ./...
go test -fuzz=FuzzInspect -fuzztime=10s -parallel=2 .
go build ./cmd/xlsxcache
```

Tests cover cache presence semantics, formula errors, shared followers, array
anchors, hidden sheets, Strict namespaces, relocated parts, invalid relationships,
malformed XML, ZIP checksums/limits, bounded reporting and CLI exit/output errors.
CI runs tests with the race detector on Linux, macOS and Windows.

Built by [Ali Farid](https://github.com/Alifarid0011). MIT licensed.
