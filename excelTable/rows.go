package excelTable

import (
	"fmt"
)

type Rows struct {
	rows [][]string
}

// TitleIndex 指定表头所在的行数
const TitleIndex = 2

// GetTitleCellIndex 查找flag在表头行中的位置
func (s *Rows) GetTitleCellIndex(str string) int {
	return findIndex(s.rows[TitleIndex], str)
}

// GetFlagRows 获取包含指定标题内容的所有行
func (s *Rows) GetFlagRows(flag, value string) ([][]string, error) {
	index := s.GetTitleCellIndex(flag)
	if index < 0 {
		return nil, fmt.Errorf("not Find Flag <%s>", flag)
	}
	var result [][]string
	for _, v := range s.rows {
		// fmt.Println("列数：", len(v))
		if len(v) <= index {
			continue
		}
		if v[index] == flag {
			result = append(result, v)
			continue
		}
		if v[index] == value {
			result = append(result, v)
		}
	}
	return result, nil
}

// CompareAndGetRows 两个表的行比较 如果含有comment,则把comment内容带入New表
func (s *Rows) CompareAndGetRows(rows [][]string, count int, newFlag string) ([][]string, error) {
	var result [][]string
	index := s.GetTitleCellIndex(newFlag)
	if index < 0 {
		return result, fmt.Errorf("not find flag %s", newFlag)
	}
	for num, _vv := range rows {
		// 新表的表头加上comment
		//error outofBounds 数组越界_vv[index] = newFlag
		if num == 0 {
			_vv = append(_vv, newFlag)
			result = append(result, _vv)
		} else {
			_vv = append(_vv, "")
		}
		// 遍历旧表内容
		for _, v := range s.rows {
			if compareStrings(v, _vv, count) {
				if _vv[index] == newFlag {
					continue
				}
				//如果旧表的row和新表的row相等，则把旧表的comments带入rows的最后
				_vv[index] = v[index]
				result = append(result, _vv)
				break
			}
		}
	}
	return result, nil
}

func (s *Rows) CompareAndWriteGetRows(rows [][]string, count int) ([][]string, error) {
	for i := 0; i < len(s.rows); i++ {
		for _, _vv := range rows {
			// 遍历旧表内容
			if compareStrings(s.rows[i], _vv, count) {
				s.rows[i] = _vv
				break
			}
		}
	}
	return s.rows, nil
}

func (s *Rows) SetValue(rows [][]string) error {
	s.rows = rows
	return nil
}
