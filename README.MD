## xlsReader
Golang Library for reading xls

## Installation

`$ go get github.com/stevenlawton/xlsReader`

## Usage

```go
package main

import (
	"fmt"
	"github.com/stevenlawton/xlsReader/xls"
	"log"
)

func main() {

	workbook, err := xls.OpenFile("small_1_sheet.xls")

	if err!=nil {
	    log.Panic(err.Error())
	}

	// Number of sheets in the workbook
	// for i := 0; i <= workbook.GetNumberSheets()-1; i++ {}

	fmt.Println(workbook.GetNumberSheets())

	sheet, err := workbook.GetSheet(0)

	if err!=nil {
	    log.Panic(err.Error())
	}

	// Print sheet name
	println(sheet.GetName())

	// Print the number of rows in the sheet
	println(sheet.GetNumberRows())

	for i := 0; i <= sheet.GetNumberRows(); i++ {
		if row, err := sheet.GetRow(i); err == nil {
			if cell, err := row.GetCol(1); err == nil {
				// Cell value, string type
				fmt.Println(cell.GetString())

				//fmt.Println(cell.GetInt64())
				//fmt.Println(cell.GetFloat64())

				// Cell type (records)
				fmt.Println(cell.GetType())

				// Receiving a formatted string, for example, for cells with a date or a percentage
				xfIndex:=cell.GetXFIndex()
				formatIndex:=workbook.GetXFbyIndex(xfIndex)
				format:=workbook.GetFormatByIndex(formatIndex.GetFormatIndex())
				fmt.Println(format.GetFormatString(cell))

			}

		}
	}
}
```
 
