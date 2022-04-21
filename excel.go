package excelstr

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

var firstCharacter = 65 // start from 'A' line

//Excel 一层结构体
func Excel(str interface{}, sheetName string) *excelize.File {
	var exce = excelize.NewFile()
	exce = WriteXlsx(exce, sheetName, firstCharacter, str)
	exce.DeleteSheet("sheet1")
	return exce
}

//Excel 多层结构体
func ExcelStruct(str interface{}) *excelize.File {
	var exce = excelize.NewFile()
	elemType := reflect.TypeOf(str)
	elemValue := reflect.ValueOf(str)
	for j := 0; j < elemType.NumField(); j++ {
		exce = WriteXlsx(exce, elemType.Field(j).Tag.Get("xlsx"), firstCharacter, elemValue.Field(j).Interface())
	}
	exce.DeleteSheet("sheet1")
	return exce
}

func WriteXlsx(exceliz *excelize.File, sheet string, firstCharacter int, records interface{}) *excelize.File {
	index := exceliz.NewSheet(sheet)
	exceliz.SetActiveSheet(index)
	t := reflect.TypeOf(records)
	s := reflect.ValueOf(records)
	if t.Kind() != reflect.Slice {
		if t.Kind() == reflect.Struct {
			for j := 0; j < t.NumField(); j++ {
				column := string(firstCharacter + j)
				exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, 1), t.Field(j).Tag.Get("xlsx"))
				exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, 2), s.Field(j).Interface())
			}
			return exceliz
		} else {
			column := string(firstCharacter)
			exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, 1), s)
		}
	} else {
		for i := 0; i < s.Len(); i++ {
			elem := s.Index(i).Interface()
			elemType := reflect.TypeOf(elem)
			elemValue := reflect.ValueOf(elem)
			if elemType.Kind() != reflect.Struct {
				column := string(firstCharacter)
				exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, i+2), elemValue)
			} else {
				for j := 0; j < elemType.NumField(); j++ {
					column := string(firstCharacter + j)
					if i == 0 {
						exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, i+1), elemType.Field(j).Tag.Get("xlsx"))
					}
					exceliz.SetCellValue(sheet, fmt.Sprintf("%s%d", column, i+2), elemValue.Field(j).Interface())
				}
			}
		}
	}
	return exceliz
}
