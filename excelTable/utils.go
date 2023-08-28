package excelTable

import (
	"log"
	"strings"

	"github.com/xuri/excelize/v2"
)

// 读入excel文件信息
func openExcel(filepath string) *excelize.File {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		if err := file.Close(); err != nil {
			log.Fatal(err)
		}
	}()
	return file
}

// 查找字符串在字符数组中的位置
func findIndex(srtArr []string, str string) int {
	upperStr := strings.ToUpper(str)
	for i, s := range srtArr {
		if strings.ToUpper(s) == upperStr {
			return i
		}
	}
	return -1
}

func compareStrings(arr1 []string, arr2 []string, count int) bool {
	if len(arr1) <= count || len(arr2) <= count {
		return false
	}
	for i := 0; i < count; i++ {
		if arr1[i] != arr2[i] {
			return false
		}
	}
	return true
}
