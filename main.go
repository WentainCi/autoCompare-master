package main

import (
	"github.com/autoCompare/excelTable"
)

func main() {
	file2name := "0425.xlsx"
	file1name := "0425_new.xlsx"
	excelTable1 := excelTable.NewExcelTable(file1name)
	excelTable2 := excelTable.NewExcelTable(file2name)
	for _, sheetVale := range excelTable1.GetSheets() {
		newExcelBook := excelTable.FinalTable(excelTable1, excelTable2, sheetVale)
		rows := excelTable1.GetRows(sheetVale)
		//将新内容替换到sheetVale的Rows中
		rows.SetValue(newExcelBook)
		//将flag和comment的样式与前一个样式相同
		excelTable1.SetColStyles(sheetVale, "flag", "comments", rows)
	}
	excelTable1.SaveToFile("new.xlsx")
}
