package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"regexp"
)

func main() {
	in := flag.String("in", "", "输入excel文件名")
	out := flag.String("out", "", "输出excel文件名")
	mSheet := flag.String("sheet", "", "excel页签")
	mCol := flag.String("col", "", "excel列")
	regexpStr := flag.String("regexp", "", "正则式")
	replaceStr := flag.String("replace", "", "替换式")

	flag.Parse()

	if *in == "" {
		flag.Usage()
		panic("必须填写参数：-in")
	}

	if *regexpStr == "" {
		panic("必须填写参数：-regexp")
	}

	if *replaceStr == "" {
		panic("必须填写参数：-replace")
	}

	file, err := excelize.OpenFile(*in)
	if err != nil {
		panic(err)
	}

	mColNum, _ := excelize.ColumnNameToNumber(*mCol)

	regexpr, err := regexp.Compile(*regexpStr)
	if err != nil {
		panic(err)
	}

	for _, sheet := range file.GetSheetList() {
		if *mSheet != "" {
			if *mSheet != sheet {
				continue
			}
		}

		rows, err := file.Rows(sheet)
		if err != nil {
			panic(err)
		}

		rowNum := 1
		for rows.Next() {
			if mColNum > 0 {
				pos, err := excelize.CoordinatesToCellName(mColNum, rowNum)
				if err != nil {
					panic(err)
				}

				cell, err := file.GetCellValue(sheet, pos)
				if err != nil {
					panic(err)
				}

				newer := regexpr.ReplaceAllString(cell, *replaceStr)
				if newer != cell {
					if err := file.SetCellStr(sheet, pos, newer); err != nil {
						panic(err)
					}
				}

			} else {
				cols, err := rows.Columns()
				if err != nil {
					panic(err)
				}

				for colNum, cell := range cols {
					colNum++

					newer := regexpr.ReplaceAllString(cell, *replaceStr)
					if newer != cell {
						pos, err := excelize.CoordinatesToCellName(colNum, rowNum)
						if err != nil {
							panic(err)
						}

						if err := file.SetCellStr(sheet, pos, newer); err != nil {
							panic(err)
						}
					}
				}
			}
			rowNum++
		}
	}

	outPath := fmt.Sprintf("%s%cnew_%s", filepath.Dir(*in), filepath.Separator, filepath.Base(*in))

	if *out != "" {
		outPath = *out
	}

	os.MkdirAll(filepath.Dir(outPath), os.ModePerm)

	if err := file.SaveAs(outPath); err != nil {
		panic(err)
	}
}
