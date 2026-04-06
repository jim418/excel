package excel

import (
    "fmt"
    "reflect"
    "strconv"
    "strings"
    "time"
    
    "github.com/xuri/excelize/v2"
)

// ========== 1. 标签定义 ==========
// 支持标签：
// - excel:"列名"           // 必填，指定Excel列名
// - excel:"列名;omitempty" // 空值时跳过（导出时）
// - format:"2006-01-02"    // 日期/时间格式
// - default:"默认值"       // 当单元格为空时使用的默认值

type FieldMeta struct {
    ColName    string        // Excel列名
    Index      int           // 列索引（导出时自动计算）
    OmitEmpty  bool          // 是否跳过空值
    DateFormat string        // 日期格式（如果有）
    Default    string        // 默认值
    FieldIdx   int           // 结构体字段索引
    FieldType  reflect.Type  // 字段类型
}

// ========== 2. 编码器（Struct → Excel） ==========
type Encoder struct {
    file       *excelize.File
    sheet      string
    startRow   int           // 起始行号（1-based）
    headers    []string      // 表头（按写入顺序）
    fieldMetas []*FieldMeta  // 字段元数据（按写入顺序）
}

// NewEncoder 创建编码器
func NewEncoder(file *excelize.File, sheet string) *Encoder {
    return &Encoder{
        file:     file,
        sheet:    sheet,
        startRow: 1,
    }
}

// SetStartRow 设置起始行（默认1，第1行通常用作表头）
func (e *Encoder) SetStartRow(row int) *Encoder {
    e.startRow = row
    return e
}

// Register 注册结构体类型，生成表头
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
        
        // 解析标签：列名;omitempty
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
        e.file.SetCellValue(e.sheet, cell, header)
    }
    return nil
}

// Encode 编码单个结构体到下一行
func (e *Encoder) Encode(obj interface{}) error {
    v := reflect.ValueOf(obj)
    if v.Kind() == reflect.Ptr {
        v = v.Elem()
    }
    if v.Kind() != reflect.Struct {
        return fmt.Errorf("expected struct, got %v", v.Kind())
    }
    
    // 计算当前行号（从表头下一行开始）
    rowNum := e.startRow + 1
    for i, meta := range e.fieldMetas {
        fieldValue := v.Field(meta.FieldIdx)
        cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
        
        // 处理空值和默认值
        if e.isEmpty(fieldValue) {
            if meta.Default != "" {
                e.file.SetCellValue(e.sheet, cell, meta.Default)
            } else if !meta.OmitEmpty {
                e.file.SetCellValue(e.sheet, cell, "")
            }
            continue
        }
        
        // 根据类型格式化值
        value := e.formatValue(fieldValue, meta)
        if err := e.file.SetCellValue(e.sheet, cell, value); err != nil {
            return fmt.Errorf("set cell %s: %w", cell, err)
        }
    }
    
    e.startRow++ // 行号递增
    return nil
}

// EncodeAll 编码多个结构体
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
        // 时间类型特殊处理：零值时间表示空
        if t, ok := v.Interface().(time.Time); ok {
            return t.IsZero()
        }
        return false
    default:
        return false
    }
}

func (e *Encoder) formatValue(v reflect.Value, meta *FieldMeta) interface{} {
    // 处理时间类型
    if t, ok := v.Interface().(time.Time); ok {
        if meta.DateFormat != "" {
            return t.Format(meta.DateFormat)
        }
        return t.Format("2006-01-02 15:04:05")
    }
    
    // 处理其他基础类型
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

// ========== 3. 解码器（Excel → Struct） ==========
type Decoder struct {
    file       *excelize.File
    sheet      string
    headerRow  int           // 表头所在行（1-based）
    colMap     map[string]int // 列名 → 列索引
    fieldMetas []*FieldMeta   // 字段元数据
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

// SetHeaderRow 设置表头行（默认第1行）
func (d *Decoder) SetHeaderRow(row int) *Decoder {
    d.headerRow = row
    return d
}

// Register 注册结构体类型，解析列映射
func (d *Decoder) Register(obj interface{}) error {
    t := reflect.TypeOf(obj)
    if t.Kind() == reflect.Ptr {
        t = t.Elem()
    }
    if t.Kind() != reflect.Struct {
        return fmt.Errorf("expected struct, got %v", t.Kind())
    }
    d.structType = t
    
    // 1. 读取Excel表头
    headers, err := d.readHeaders()
    if err != nil {
        return err
    }
    
    // 2. 建立列名→索引映射
    for i, h := range headers {
        d.colMap[strings.TrimSpace(h)] = i
    }
    
    // 3. 解析结构体标签，匹配列
    for i := 0; i < t.NumField(); i++ {
        field := t.Field(i)
        tag := field.Tag.Get("excel")
        if tag == "" || tag == "-" {
            continue
        }
        
        // 解析标签：列名;omitempty
        parts := strings.Split(tag, ";")
        colName := parts[0]
        
        colIdx, ok := d.colMap[colName]
        if !ok {
            continue // 列不存在时跳过（允许部分匹配）
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

// readHeaders 读取表头行
func (d *Decoder) readHeaders() ([]string, error) {
    rows, err := d.file.Rows(d.sheet)
    if err != nil {
        return nil, err
    }
    defer rows.Close()
    
    // 跳到表头行
    currentRow := 1
    for rows.Next() {
        if currentRow == d.headerRow {
            return rows.Columns()
        }
        currentRow++
    }
    return nil, fmt.Errorf("header row %d not found", d.headerRow)
}

// Decode 解码下一行到结构体
func (d *Decoder) Decode(obj interface{}) (bool, error) {
    // 简化实现：这里需要维护行迭代器状态
    // 完整实现需要缓存rows对象，这里展示核心逻辑
    
    v := reflect.ValueOf(obj)
    if v.Kind() != reflect.Ptr {
        return false, fmt.Errorf("expected pointer to struct")
    }
    v = v.Elem()
    if v.Kind() != reflect.Struct {
        return false, fmt.Errorf("expected struct, got %v", v.Kind())
    }
    
    // 实际使用时需要从Excel读取一行数据
    // 这里展示字段赋值的核心逻辑
    for _, meta := range d.fieldMetas {
        field := v.Field(meta.FieldIdx)
        if !field.CanSet() {
            continue
        }
        
        // 从Excel单元格读取值（示例中简化）
        cellValue := "示例值" // 实际应从当前行获取
        
        if err := d.setFieldValue(field, cellValue, meta); err != nil {
            return false, fmt.Errorf("set field %s: %w", meta.ColName, err)
        }
    }
    
    return true, nil
}

// setFieldValue 将Excel单元格值设置到结构体字段
func (d *Decoder) setFieldValue(field reflect.Value, cellValue string, meta *FieldMeta) error {
    if cellValue == "" && meta.Default != "" {
        cellValue = meta.Default
    }
    if cellValue == "" {
        return nil
    }
    
    // 处理时间类型
    if meta.DateFormat != "" {
        t, err := time.Parse(meta.DateFormat, cellValue)
        if err != nil {
            return fmt.Errorf("parse date %q with format %q: %w", cellValue, meta.DateFormat, err)
        }
        field.Set(reflect.ValueOf(t))
        return nil
    }
    
    // 处理基础类型
    switch field.Kind() {
    case reflect.String:
        field.SetString(cellValue)
    case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
        val, err := strconv.ParseInt(cellValue, 10, 64)
        if err != nil {
            return err
        }
        field.SetInt(val)
    case reflect.Float32, reflect.Float64:
        val, err := strconv.ParseFloat(cellValue, 64)
        if err != nil {
            return err
        }
        field.SetFloat(val)
    case reflect.Bool:
        val, err := strconv.ParseBool(cellValue)
        if err != nil {
            return err
        }
        field.SetBool(val)
    default:
        return fmt.Errorf("unsupported field type: %v", field.Kind())
    }
    return nil
}
