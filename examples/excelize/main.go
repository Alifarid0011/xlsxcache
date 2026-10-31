package main

import (
	"github.com/xuri/excelize/v2"
	"log"
)

func main() {
	f := excelize.NewFile()
	defer f.Close()
	check := func(err error) {
		if err != nil {
			log.Fatal(err)
		}
	}
	check(f.SetSheetName("Sheet1", "Sales"))
	check(f.SetSheetRow("Sales", "A1", &[]any{"region", "net", "tax", "total"}))
	check(f.SetSheetRow("Sales", "A2", &[]any{"North", 100}))
	check(f.SetCellFormula("Sales", "C2", "B2*0.2"))
	check(f.SetCellFormula("Sales", "D2", "B2+C2"))
	check(f.SaveAs("missing-cache.xlsx"))
}
