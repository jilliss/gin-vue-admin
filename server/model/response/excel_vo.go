package response

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"strconv"
)

type ExcelDo struct {
	Name string `excel:"name"`
	Age  int    `excel:"age" check:"age"`
	Sex  int    `excel:"sex" dic:"sex"`
}

func WriteExcleExample(excelDos []*ExcelDo) {
	xlsx := excelize.NewFile()
	index := xlsx.NewSheet("Sheet1")
	for rowIndex, excelDo := range excelDos {
		d := reflect.TypeOf(excelDo).Elem()
		for fieldIndex := 0; fieldIndex < d.NumField(); fieldIndex++ {
			// 表头
			if rowIndex == 0 {
				column := strconv.Itoa(fieldIndex + int('A'))
				name := d.Field(fieldIndex).Tag.Get("excel")
				err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, rowIndex+1), name)
				if err != nil {
					return
				}
			}
			// 字符溢问题 TODO
			column := strconv.Itoa(fieldIndex + int('A'))
			switch d.Field(fieldIndex).Type.String() {
			case "string":
				err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, rowIndex+2), reflect.ValueOf(d).Elem().Field(fieldIndex).String())
				if err != nil {
					return
				}
			case "int32", "int", "int64":
				// 字典转换 TODO
				dicKye := d.Field(fieldIndex).Tag.Get("dic")
				cKye := d.Field(fieldIndex).Tag.Get("check")
				fmt.Println(dicKye, cKye)
				err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, rowIndex+2), reflect.ValueOf(d).Elem().Field(fieldIndex).Int())
				if err != nil {
					return
				}
			case "bool":
				err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, rowIndex+2), reflect.ValueOf(d).Elem().Field(fieldIndex).Bool())
				if err != nil {
					return
				}
			case "float32", "float64":
				err := xlsx.SetCellValue("Sheet1", fmt.Sprintf("%s%d", column, rowIndex+2), reflect.ValueOf(d).Elem().Field(fieldIndex).Float())
				if err != nil {
					return
				}
			}
		}
	}
	xlsx.SetActiveSheet(index)
	err := xlsx.SaveAs("auto.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
