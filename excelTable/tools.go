package excelTable

import (
	"encoding/json"
	"fmt"
	"log"
)

// FinalTable 相同的sheet进行比较
func FinalTable(tableNew *ExcelTable, tableOld *ExcelTable, sheet string) [][]string {
	rows1 := tableNew.GetRows(sheet)
	rows2 := tableOld.GetRows(sheet)
	_ = rows2
	//找出符合条件的所有行数据
	OldRows, err := rows1.GetFlagRows("flag", "old")
	if err != nil {
		log.Println(err)
	}
	//fmt.Printf("%+v\n", OldRows)
	index := rows1.GetTitleCellIndex("flag")
	result, er := rows2.CompareAndGetRows(OldRows, index, "comments")
	if er != nil {
		log.Println(er)
	}
	fmt.Println("======================")
	for k, v := range result {
		b, _ := json.Marshal(v)
		fmt.Printf("%d %v\n", k, string(b))
	}
	result, _ = rows1.CompareAndWriteGetRows(result, index)
	fmt.Println("====All==================")
	for k, v := range result {
		b, _ := json.Marshal(v)
		fmt.Printf("%d %v\n", k, string(b))
	}
	return result
}
