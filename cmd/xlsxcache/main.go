package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"runtime/debug"

	xc "github.com/Alifarid0011/xlsxcache"
)

var version = "dev"

const usage = `Usage: xlsxcache [--json] workbook.xlsx

Check saved formula results without recalculating or changing the workbook.
Checks explicit formula cells on all worksheets, including hidden sheets.
A clean report does not prove freshness or correctness.

Flags must precede the path. Use -- before a path beginning with a dash.
  --json       print JSON (schema_version 1)
  --version    print version
  --help       print this help

Exit codes: 0 no findings; 1 missing caches or cached errors; 2 input/output error.
`

func main() { os.Exit(run(os.Args[1:], os.Stdout, os.Stderr)) }

func run(args []string, stdout, stderr io.Writer) int {
	fail := func(err error) int { fmt.Fprintf(stderr, "xlsxcache: %v\n", err); return 2 }
	flags := flag.NewFlagSet("xlsxcache", flag.ContinueOnError)
	flags.SetOutput(stderr)
	jsonOutput := flags.Bool("json", false, "print JSON")
	showVersion := flags.Bool("version", false, "print version")
	help := flags.Bool("help", false, "print help")
	flags.Usage = func() { fmt.Fprint(stderr, usage) }
	if err := flags.Parse(args); err != nil {
		if errors.Is(err, flag.ErrHelp) {
			return writeOutput(stdout, stderr, []byte(usage))
		}
		return 2
	}
	if *help {
		return writeOutput(stdout, stderr, []byte(usage))
	}
	if *showVersion {
		v := version
		if v == "dev" {
			if info, ok := debug.ReadBuildInfo(); ok && info.Main.Version != "" && info.Main.Version != "(devel)" {
				v = info.Main.Version
			}
		}
		return writeOutput(stdout, stderr, []byte("xlsxcache "+v+"\n"))
	}
	if flags.NArg() != 1 {
		return fail(fmt.Errorf("expected one workbook path; use --help for usage"))
	}
	filename := flags.Arg(0)
	f, err := os.Open(filename)
	if err != nil {
		return fail(err)
	}
	info, err := f.Stat()
	if err != nil {
		f.Close()
		return fail(err)
	}
	if !info.Mode().IsRegular() {
		f.Close()
		return fail(fmt.Errorf("%q is not a regular file", filename))
	}
	report, inspectErr := xc.Inspect(f, info.Size())
	closeErr := f.Close()
	if inspectErr != nil {
		return fail(fmt.Errorf("%q: %w", filename, inspectErr))
	}
	if closeErr != nil {
		return fail(closeErr)
	}
	var output bytes.Buffer
	if *jsonOutput {
		doc := struct {
			SchemaVersion int    `json:"schema_version"`
			File          string `json:"file"`
			Freshness     string `json:"freshness"`
			xc.Report
		}{1, filename, "not_checked", report}
		encoder := json.NewEncoder(&output)
		encoder.SetIndent("", "  ")
		if err := encoder.Encode(doc); err != nil {
			return fail(err)
		}
	} else {
		fmt.Fprintf(&output, "%q: %d worksheets, %d formula cells, %d missing caches, %d cached errors\n", filename, report.Sheets, report.Formulas, report.Missing, report.Errors)
		for _, finding := range report.Findings {
			fmt.Fprintf(&output, "  %q!%s: %s\n", finding.Sheet, finding.Cell, finding.Code)
		}
		if report.Truncated {
			fmt.Fprintln(&output, "  locations truncated to the first 1000; counts include every finding")
		}
		fmt.Fprintln(&output, "Cache freshness not checked; saved results may be stale.")
	}
	if code := writeOutput(stdout, stderr, output.Bytes()); code != 0 {
		return code
	}
	if report.Missing+report.Errors > 0 {
		return 1
	}
	return 0
}
func writeOutput(stdout, stderr io.Writer, data []byte) int {
	n, err := stdout.Write(data)
	if err == nil && n != len(data) {
		err = io.ErrShortWrite
	}
	if err != nil {
		fmt.Fprintf(stderr, "xlsxcache: write output: %v\n", err)
		return 2
	}
	return 0
}
