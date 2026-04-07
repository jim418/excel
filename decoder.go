package excel

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ========== 基础解码器 ==========

// Decoder 基础解码器
type Decoder struct {
	file       *excelize.File
	sheet      string
	headerRow  int
	colMap     map[string]int
	fieldMetas []*FieldMeta
	structType reflect.Type
}

// NewDecoder 创建解码器
func NewDecoder(file *excelize.File, sheet string) *Decoder {
	return &Decoder{
		file:      file,
		sheet:     sheet,
		headerRow: 1,
		colMap:    make(map[string]int),
	}
}

// SetHeaderRow 设置表头行
func (d *Decoder) SetHeaderRow(row int) *Decoder {
	d.headerRow = row
	return d
}

// Register 注册结构体类型
func (d *Decoder) Register(obj interface{}) error {
	t := reflect.TypeOf(obj)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return fmt.Errorf("expected struct, got %v", t.Kind())
	}
	d.structType = t

	// 读取表头
	headers, err := d.readHeaders()
	if err != nil {
		return err
	}

	for i, h := range headers {
		d.colMap[strings.TrimSpace(h)] = i
	}

	// 解析结构体字段
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		tag := field.Tag.Get("excel")
		if tag == "" || tag == "-" {
			continue
		}

		colName := strings.Split(tag, ";")[0]
		colIdx, ok := d.colMap[colName]
		if !ok {
			continue
		}

		meta := &FieldMeta{
			ColName:    colName,
			Index:      colIdx,
			DateFormat: field.Tag.Get("format"),
			Default:    field.Tag.Get("default"),
			FieldIdx:   i,
			FieldType:  field.Type,
		}
		d.fieldMetas = append(d.fieldMetas, meta)
	}

	return nil
}

// DecodeAll 解码所有行
func (d *Decoder) DecodeAll(result interface{}) error {
	resultPtr := reflect.ValueOf(result)
	if resultPtr.Kind() != reflect.Ptr || resultPtr.Elem().Kind() != reflect.Slice {
		return fmt.Errorf("result must be pointer to slice")
	}
	slicePtr := resultPtr.Elem()
	elemType := slicePtr.Type().Elem()

	rows, err := d.file.Rows(d.sheet)
	if err != nil {
		return err
	}
	defer rows.Close()

	// 跳到表头
	for i := 1; i < d.headerRow; i++ {
		if !rows.Next() {
			return fmt.Errorf("not enough rows for header")
		}
	}
	rows.Next() // 跳过头行

	rowNum := d.headerRow + 1
	for rows.Next() {
		cols, _ := rows.Columns()

		elemPtr := reflect.New(elemType)
		elem := elemPtr.Elem()
		if elemType.Kind() == reflect.Ptr {
			elem = elemPtr
			elemPtr = reflect.New(elemType.Elem())
			elem = elemPtr.Elem()
		}

		for _, meta := range d.fieldMetas {
			field := elem.Field(meta.FieldIdx)
			if !field.CanSet() {
				continue
			}

			var cellValue string
			if meta.Index < len(cols) {
				cellValue = cols[meta.Index]
			}

			d.setFieldValue(field, cellValue, meta, rowNum)
		}

		if elemType.Kind() == reflect.Ptr {
			slicePtr.Set(reflect.Append(slicePtr, elemPtr))
		} else {
			slicePtr.Set(reflect.Append(slicePtr, elem))
		}
		rowNum++
	}

	return nil
}

func (d *Decoder) readHeaders() ([]string, error) {
	rows, err := d.file.Rows(d.sheet)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	currentRow := 1
	for rows.Next() {
		if currentRow == d.headerRow {
			return rows.Columns()
		}
		currentRow++
	}
	return nil, fmt.Errorf("header row %d not found", d.headerRow)
}

func (d *Decoder) setFieldValue(field reflect.Value, cellValue string, meta *FieldMeta, rowNum int) {
	if cellValue == "" && meta.Default != "" {
		cellValue = meta.Default
	}
	if cellValue == "" {
		return
	}

	// 时间类型
	if meta.DateFormat != "" || meta.FieldType == reflect.TypeOf(time.Time{}) {
		if t, ok := TryParseExcelDate(cellValue, meta.DateFormat, false, 0); ok {
			field.Set(reflect.ValueOf(t))
		}
		return
	}

	// 基础类型
	switch field.Kind() {
	case reflect.String:
		field.SetString(cellValue)
	case reflect.Int, reflect.Int64:
		if val, err := strconv.ParseInt(cellValue, 10, 64); err == nil {
			field.SetInt(val)
		}
	case reflect.Float64:
		if val, err := strconv.ParseFloat(cellValue, 64); err == nil {
			field.SetFloat(val)
		}
	case reflect.Bool:
		if val, err := strconv.ParseBool(cellValue); err == nil {
			field.SetBool(val)
		}
	}
}
