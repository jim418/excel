package excel

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ========== 流式编码器（大数据量导出） ==========

// StreamEncoder 流式编码器
type StreamEncoder struct {
	file       *excelize.File
	sheet      string
	stream     *excelize.StreamWriter
	startRow   int
	headers    []string
	fieldMetas []*FieldMeta
	curRow     int
	use1904    bool
}

// NewStreamEncoder 创建流式编码器
func NewStreamEncoder(file *excelize.File, sheet string) (*StreamEncoder, error) {
	stream, err := file.NewStreamWriter(sheet)
	if err != nil {
		return nil, err
	}

	return &StreamEncoder{
		file:     file,
		sheet:    sheet,
		stream:   stream,
		startRow: 1,
		curRow:   1,
		use1904:  false,
	}, nil
}

// SetStartRow 设置起始行
func (e *StreamEncoder) SetStartRow(row int) *StreamEncoder {
	e.startRow = row
	e.curRow = row
	return e
}

// SetUse1904 设置使用1904日期系统
func (e *StreamEncoder) SetUse1904(use1904 bool) *StreamEncoder {
	e.use1904 = use1904
	return e
}

// Register 注册结构体类型
func (e *StreamEncoder) Register(obj interface{}) error {
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
func (e *StreamEncoder) WriteHeaders() error {
	row := make([]interface{}, len(e.headers))
	for i, header := range e.headers {
		row[i] = header
	}
	cell, _ := excelize.CoordinatesToCellName(1, e.curRow)
	if err := e.stream.SetRow(cell, row); err != nil {
		return err
	}
	e.curRow++
	return nil
}

// Encode 编码单个结构体
func (e *StreamEncoder) Encode(obj interface{}) error {
	v := reflect.ValueOf(obj)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}

	row := make([]interface{}, len(e.fieldMetas))
	for i, meta := range e.fieldMetas {
		fieldValue := v.Field(meta.FieldIdx)
		row[i] = e.formatValueForStream(fieldValue, meta)
	}

	cell, _ := excelize.CoordinatesToCellName(1, e.curRow)
	if err := e.stream.SetRow(cell, row); err != nil {
		return err
	}
	e.curRow++
	return nil
}

// EncodeAll 批量编码
func (e *StreamEncoder) EncodeAll(objs interface{}) error {
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

// Flush 刷新缓冲区
func (e *StreamEncoder) Flush() error {
	return e.stream.Flush()
}

func (e *StreamEncoder) formatValueForStream(v reflect.Value, meta *FieldMeta) interface{} {
	if t, ok := v.Interface().(time.Time); ok {
		if t.IsZero() {
			return nil
		}
		if meta.DateFormat != "" && meta.DateFormat != "excel" {
			return t.Format(meta.DateFormat)
		}
		return TimeToExcelTime(t, e.use1904)
	}

	if e.isEmptyForStream(v) {
		if meta.Default != "" {
			return meta.Default
		}
		if meta.OmitEmpty {
			return nil
		}
		return ""
	}

	return v.Interface()
}

func (e *StreamEncoder) isEmptyForStream(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.String:
		return v.String() == ""
	case reflect.Int, reflect.Int64:
		return v.Int() == 0
	case reflect.Float64:
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
