package excel

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ========== 基础编码器 ==========

// FieldMeta 字段元数据
type FieldMeta struct {
	ColName    string       // Excel列名
	Index      int          // 列索引
	OmitEmpty  bool         // 是否跳过空值
	DateFormat string       // 日期格式
	Default    string       // 默认值
	FieldIdx   int          // 结构体字段索引
	FieldType  reflect.Type // 字段类型
	Validation *ValidationRule // 校验规则（可选）
}

// Encoder 基础编码器
type Encoder struct {
	file       *excelize.File
	sheet      string
	startRow   int
	headers    []string
	fieldMetas []*FieldMeta
}

// NewEncoder 创建编码器
func NewEncoder(file *excelize.File, sheet string) *Encoder {
	return &Encoder{
		file:     file,
		sheet:    sheet,
		startRow: 1,
	}
}

// SetStartRow 设置起始行
func (e *Encoder) SetStartRow(row int) *Encoder {
	e.startRow = row
	return e
}

// Register 注册结构体类型
func (e *Encoder) Register(obj interface{}) error {
	t := reflect.TypeOf(obj)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return fmt.Errorf("expected struct, got %v", t.Kind())
	}

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		tag := field.Tag.Get("excel")
		if tag == "" || tag == "-" {
			continue
		}

		parts := strings.Split(tag, ";")
		colName := parts[0]
		omitEmpty := false
		for _, p := range parts[1:] {
			if p == "omitempty" {
				omitEmpty = true
			}
		}

		meta := &FieldMeta{
			ColName:    colName,
			OmitEmpty:  omitEmpty,
			DateFormat: field.Tag.Get("format"),
			Default:    field.Tag.Get("default"),
			FieldIdx:   i,
			FieldType:  field.Type,
		}
		e.fieldMetas = append(e.fieldMetas, meta)
		e.headers = append(e.headers, colName)
	}

	return nil
}

// WriteHeaders 写入表头
func (e *Encoder) WriteHeaders() error {
	for i, header := range e.headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, e.startRow)
		if err := e.file.SetCellValue(e.sheet, cell, header); err != nil {
			return err
		}
	}
	return nil
}

// Encode 编码单个结构体
func (e *Encoder) Encode(obj interface{}) error {
	v := reflect.ValueOf(obj)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}
	if v.Kind() != reflect.Struct {
		return fmt.Errorf("expected struct, got %v", v.Kind())
	}

	rowNum := e.startRow + 1
	for i, meta := range e.fieldMetas {
		fieldValue := v.Field(meta.FieldIdx)
		cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)

		// 处理空值
		if e.isEmpty(fieldValue) {
			if meta.Default != "" {
				e.file.SetCellValue(e.sheet, cell, meta.Default)
			} else if !meta.OmitEmpty {
				e.file.SetCellValue(e.sheet, cell, "")
			}
			continue
		}

		value := e.formatValue(fieldValue, meta)
		if err := e.file.SetCellValue(e.sheet, cell, value); err != nil {
			return fmt.Errorf("set cell %s: %w", cell, err)
		}
	}

	e.startRow = rowNum
	return nil
}

// EncodeAll 批量编码
func (e *Encoder) EncodeAll(objs interface{}) error {
	v := reflect.ValueOf(objs)
	if v.Kind() != reflect.Slice {
		return fmt.Errorf("expected slice, got %v", v.Kind())
	}

	for i := 0; i < v.Len(); i++ {
		if err := e.Encode(v.Index(i).Interface()); err != nil {
			return fmt.Errorf("encode item %d: %w", i, err)
		}
	}
	return nil
}

func (e *Encoder) isEmpty(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.String:
		return v.String() == ""
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int() == 0
	case reflect.Float32, reflect.Float64:
		return v.Float() == 0
	case reflect.Struct:
		if t, ok := v.Interface().(time.Time); ok {
			return t.IsZero()
		}
		return false
	default:
		return false
	}
}

func (e *Encoder) formatValue(v reflect.Value, meta *FieldMeta) interface{} {
	if t, ok := v.Interface().(time.Time); ok {
		if meta.DateFormat != "" {
			return t.Format(meta.DateFormat)
		}
		return t.Format("2006-01-02 15:04:05")
	}

	switch v.Kind() {
	case reflect.String:
		return v.String()
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int()
	case reflect.Float32, reflect.Float64:
		return v.Float()
	case reflect.Bool:
		return v.Bool()
	default:
		return fmt.Sprintf("%v", v.Interface())
	}
}
