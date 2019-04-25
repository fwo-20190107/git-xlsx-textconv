package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	xls "github.com/fwo-20190107/xls"
	xlsx "github.com/tealeg/xlsx"
)

func main() {
	if len(os.Args) != 2 {
		log.Fatal("Usage: git-xlsx-textconv file.xls[x]")
	}
	excelFileName := os.Args[1]
	pos := strings.LastIndex(excelFileName, ".")
	ext := excelFileName[pos:]

	if ext == ".xls" {
		xlFile, err := xls.Open(excelFileName, "utf-8")
		if err != nil {
			log.Fatal(err)
		}
		for i := 0; i < xlFile.NumSheets(); i++ {
			if sheet := xlFile.GetSheet(i); sheet != nil {
				var r uint16
				for r = 0; r < sheet.MaxRow; r++ {
					row := sheet.Row(int(r))

					cels := make([]string, 3)
					arr := []int{1, 2, 3}
					for n, _ := range arr {
						var s string
						s = row.Col(0)

						s = strings.Replace(s, "\\", "\\\\", -1)
						s = strings.Replace(s, "\n", "\\n", -1)
						s = strings.Replace(s, "\r", "\\r", -1)
						s = strings.Replace(s, "\t", "\\t", -1)

						cels[n] = s
					}
					fmt.Printf("[%s] %s\n", sheet.Name, strings.Join(cels, "\t"))
				}
			}
		}
	}
	if ext == ".xlsx" {
		xlFile, err := xlsx.OpenFile(excelFileName)
		if err != nil {
			log.Fatal(err)
		}

		for _, sheet := range xlFile.Sheets {
			for _, row := range sheet.Rows {
				if row == nil {
					continue
				}
				cels := make([]string, len(row.Cells))
				for i, cell := range row.Cells {
					var s string
					if cell.Type() == xlsx.CellTypeStringFormula {
						s = cell.Formula()
					} else {
						s = cell.String()
					}

					s = strings.Replace(s, "\\", "\\\\", -1)
					s = strings.Replace(s, "\n", "\\n", -1)
					s = strings.Replace(s, "\r", "\\r", -1)
					s = strings.Replace(s, "\t", "\\t", -1)

					cels[i] = s
				}
				fmt.Printf("[%s] %s\n", sheet.Name, strings.Join(cels, "\t"))
			}
		}
	}
}
