package excelTable

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type ExcelTable struct {
	FileName string
	file     *excelize.File
	rows     map[string]*Rows
}

// NewExcelTable 获取Excel为一个对象
func NewExcelTable(fileName string) *ExcelTable {
	result := &ExcelTable{
		FileName: fileName,
		file:     openExcel(fileName),
		rows:     make(map[string]*Rows),
	}
	return result
}

// GetRows 传入sheet，获取它的每一行
func (s *ExcelTable) GetRows(sheet string) *Rows {
	if rows, ok := s.rows[sheet]; ok {
		return rows
	} else {
		_rows, _ := s.file.GetRows(sheet)
		rows = &Rows{rows: _rows}
		s.rows[sheet] = rows
		return rows
	}
}

// GetSheets 获取Excel表中的 sheets名称到数组
func (s *ExcelTable) GetSheets() []string {
	return s.file.GetSheetList()
}

// SaveToFile 保存Excel表到文件
func (s *ExcelTable) SaveToFile(fileName string) {

	for sheet, row := range s.rows {
		metrics, err := s.file.GetMergeCells(sheet)
		if err != nil {
			fmt.Println(err)
		}
		for i, v := range row.rows {
			params := strSliceToInterfaseSlice(v)
			if err := s.file.SetSheetRow(sheet, fmt.Sprintf("a%d", i+1), &params); err != nil {
				fmt.Println(err)
			}
		}

		{
			cols, err := s.file.GetCols(sheet)
			fmt.Println(cols, err)
		}
		fmt.Println(metrics)
	}
	s.file.SaveAs(fileName)
}

func strSliceToInterfaseSlice(strs []string) []interface{} {
	var result []interface{}
	for _, v := range strs {
		result = append(result, v)
	}
	return result
}

// 设置样式
func (s *ExcelTable) SetColStyles(sheet, start, end string, rows *Rows) error {
	//遍历表单，对每个单元格的样式进行复制
	startIndex := rows.GetTitleCellIndex(start) //9
	endIndex := rows.GetTitleCellIndex(end)     //10
	if startIndex < 0 {
		return nil
	}
	//没有comments的情况
	if endIndex < 0 {
		endIndex = startIndex
	}
	//返回要复制样式的单元格列名
	colName, _ := excelize.ColumnNumberToName(startIndex) //I
	for i := 0; i < len(rows.rows); i++ {
		//单元格坐标
		axis := colName + strconv.Itoa(i+1)
		//获取单元格样式索引
		style, err := s.file.GetCellStyle(sheet, axis)
		if err != nil {
			return err
		}
		//给需要样式的单元格附上样式
		for j := startIndex; j <= endIndex; j++ {
			//获取单元格坐标
			colName2, _ := excelize.ColumnNumberToName(j + 1)
			axis2 := colName2 + strconv.Itoa(i+1)
			err = s.file.SetCellStyle(sheet, axis2, axis2, style)
			if err != nil {
				return err
			}
		}
	}
	return nil
}
