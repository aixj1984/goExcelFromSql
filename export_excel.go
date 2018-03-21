package main

import (
	"fmt"
	//	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/astaxie/beego/orm"
)

func ExportExcel(result *[]orm.Params, filepath string, column_map map[string]string) bool {
	//xlsx := excelize.CreateFile()

	columns := GetColumns(result)

	if columns == nil {
		return false
	}

	xlsx := excelize.NewFile()

	//Set value of a cell.

	//fmt.Println(result)
	//fmt.Println(columns)
	//categories := map[string]string{"A1": "Small", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	//values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}

	categories := make(map[string]string, 0)

	for k, v := range *columns {
		key := precessCategories(k)
		categories[key+"1"] = v

	}
	//fmt.Println(categories)
	for k, v := range categories {
		if _, ok := column_map[v]; ok {
			xlsx.SetCellValue("Sheet1", k, column_map[v])
		} else {
			xlsx.SetCellValue("Sheet1", k, v)
		}
	}

	values := make(map[string]string, 0)

	for k1, v1 := range *result {
		c := 0
		for k2, v2 := range v1 {
			var value string
			i := getArrKey(columns, k2)
			key := precessCategories(i) + strconv.Itoa(k1+2)
			if v2 == nil {
				value = "NULL"
			} else {
				if _, ok := v2.(string); ok {
					value = v2.(string)
				} else {
					value = ""
				}
			}
			values[key] = value
			//fmt.Println(key)
			c++
		}

	}
	//fmt.Println(values)
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}

	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(2)
	// Save xlsx file by the given path.
	err := xlsx.SaveAs(filepath)
	if err != nil {
		fmt.Println(err)
		return false
	}
	return true
}

func GetColumns(values *[]orm.Params) *[]string {
	if len(*values) > 0 {
		var columns []string
		columns = make([]string, len((*values)[0]))
		var index = 0
		for k, _ := range (*values)[0] {
			//fmt.Println(k, v)
			columns[index] = k
			index++
		}
		return &columns
	}
	return nil
}

//map is not ordered
func getArrKey(arr *[]string, value string) int {
	for k, v := range *arr {
		if v == value {
			return k
		}
	}
	return -1
}

//excel
func precessCategories(k int) string {
	az := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if k < 26 {
		return string(az[k])
	} else {
		k1 := int((k + 1) / 26)
		k2 := (k + 1) % 26
		return string(az[k1]) + string(az[k2])
	}
}

func checkErr(err error) {
	if err != nil {
		panic(err)
	}
}
