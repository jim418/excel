package excel

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ========== Excel原生数据验证 ==========

// ExcelDataValidation Excel数据验证配置
type ExcelDataValidation struct {
	Type         string // "whole", "decimal", "list", "date", "time", "textLength", "custom"
	Operator     string // "between", "notBetween", "equal", "notEqual", "greaterThan", ...
	Formula1     string
	Formula2     string
	AllowBlank   bool
	ShowDropDown bool
	ErrorTitle   string
	ErrorMsg     string
	ErrorStyle   string // "stop", "warning", "information"
	InputTitle   string
	InputMsg     string
}

// AddDataValidationToColumn 为指定列添加数据验证
func AddDataValidationToColumn(f *excelize.File, sheet string, column string, startRow, endRow int, dv *ExcelDataValidation) error {
	sqref := fmt.Sprintf("%s%d:%s%d", column, startRow, column, endRow)

	dataVal := excelize.NewDataValidation(dv.AllowBlank)
	dataVal.SetSqref(sqref)

	if dv.ShowDropDown && dv.Type == "list" {
		dataVal.SetDropList(strings.Split(dv.Formula1, ","))
	} else {
		switch dv.Type {
		case "whole", "decimal", "date", "time", "textLength":
			var op excelize.DataValidationOperator
			switch dv.Operator {
			case "between":
				op = excelize.DataValidationOperatorBetween
			case "notBetween":
				op = excelize.DataValidationOperatorNotBetween
			case "equal":
				op = excelize.DataValidationOperatorEqual
			case "notEqual":
				op = excelize.DataValidationOperatorNotEqual
			case "greaterThan":
				op = excelize.DataValidationOperatorGreaterThan
			case "lessThan":
				op = excelize.DataValidationOperatorLessThan
			case "greaterThanOrEqual":
				op = excelize.DataValidationOperatorGreaterThanOrEqual
			case "lessThanOrEqual":
				op = excelize.DataValidationOperatorLessThanOrEqual
			default:
				op = excelize.DataValidationOperatorBetween
			}

			var valType excelize.DataValidationType
			switch dv.Type {
			case "whole":
				valType = excelize.DataValidationTypeWhole
			case "decimal":
				valType = excelize.DataValidationTypeDecimal
			case "date":
				valType = excelize.DataValidationTypeDate
			case "time":
				valType = excelize.DataValidationTypeTime
			case "textLength":
				valType = excelize.DataValidationTypeTextLength
			default:
				valType = excelize.DataValidationTypeWhole
			}

			f1, _ := strconv.ParseFloat(dv.Formula1, 64)
			f2, _ := strconv.ParseFloat(dv.Formula2, 64)
			dataVal.SetRange(int(f1), int(f2), valType, op)

		case "list":
			dataVal.SetSqrefDropList(dv.Formula1)
		}
	}

	if dv.ErrorTitle != "" || dv.ErrorMsg != "" {
		var style excelize.DataValidationErrorStyle
		switch dv.ErrorStyle {
		case "stop":
			style = excelize.DataValidationErrorStyleStop
		case "warning":
			style = excelize.DataValidationErrorStyleWarning
		case "information":
			style = excelize.DataValidationErrorStyleInformation
		default:
			style = excelize.DataValidationErrorStyleStop
		}
		dataVal.SetError(style, dv.ErrorTitle, dv.ErrorMsg)
	}

	if dv.InputTitle != "" || dv.InputMsg != "" {
		dataVal.SetInput(dv.InputTitle, dv.InputMsg)
	}

	return f.AddDataValidation(sheet, dataVal)
}

// AddDropdownList 添加下拉列表
func AddDropdownList(f *excelize.File, sheet string, column string, startRow, endRow int, options []string, allowBlank bool) error {
	dv := &ExcelDataValidation{
		Type:         "list",
		Formula1:     strings.Join(options, ","),
		AllowBlank:   allowBlank,
		ShowDropDown: true,
		ErrorTitle:   "无效输入",
		ErrorMsg:     "请从下拉列表中选择有效值",
		ErrorStyle:   "stop",
	}
	return AddDataValidationToColumn(f, sheet, column, startRow, endRow, dv)
}

// AddRangeValidation 添加数值范围验证
func AddRangeValidation(f *excelize.File, sheet string, column string, startRow, endRow int, minVal, maxVal float64, allowBlank bool) error {
	dv := &ExcelDataValidation{
		Type:       "whole",
		Operator:   "between",
		Formula1:   fmt.Sprintf("%f", minVal),
		Formula2:   fmt.Sprintf("%f", maxVal),
		AllowBlank: allowBlank,
		ErrorTitle: "数值超出范围",
		ErrorMsg:   fmt.Sprintf("请输入 %.2f 到 %.2f 之间的数值", minVal, maxVal),
		ErrorStyle: "stop",
	}
	return AddDataValidationToColumn(f, sheet, column, startRow, endRow, dv)
}

// AddDateRangeValidation 添加日期范围验证
func AddDateRangeValidation(f *excelize.File, sheet string, column string, startRow, endRow int, minDate, maxDate string, allowBlank bool) error {
	dv := &ExcelDataValidation{
		Type:       "date",
		Operator:   "between",
		Formula1:   minDate,
		Formula2:   maxDate,
		AllowBlank: allowBlank,
		ErrorTitle: "日期超出范围",
		ErrorMsg:   fmt.Sprintf("请输入 %s 到 %s 之间的日期", minDate, maxDate),
		ErrorStyle: "stop",
	}
	return AddDataValidationToColumn(f, sheet, column, startRow, endRow, dv)
}
