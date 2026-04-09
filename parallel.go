package excel

import (
	"fmt"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

// ========== 并行解码器 ==========

// ParallelDecoder 并行解码器
type ParallelDecoder struct {
	file       *excelize.File
	sheet      string
	headerRow  int
	workers    int
	batchSize  int
	collector  *ErrorCollector
	use1904    bool
	structType reflect.Type
	fieldMetas []*FieldMeta
	colMap     map[string]int
}

// NewParallelDecoder 创建并行解码器
func NewParallelDecoder(file *excelize.File, sheet string) *ParallelDecoder {
	return &ParallelDecoder{
		file:       file,
		sheet:      sheet,
		headerRow:  1,
		workers:    4,
		batchSize:  1000,
		collector:  NewErrorCollector(),
		use1904:    false,
		colMap:     make(map[string]int),
	}
}

// SetWorkers 设置并发数
func (d *ParallelDecoder) SetWorkers(n int) *ParallelDecoder {
	d.workers = n
	return d
}

// SetBatchSize 设置批次大小
func (d *ParallelDecoder) SetBatchSize(size int) *ParallelDecoder {
	d.batchSize = size
	return d
}

// SetCollector 设置错误收集器
func (d *ParallelDecoder) SetCollector(c *ErrorCollector) *ParallelDecoder {
	d.collector = c
	return d
}

// Register 注册结构体类型
func (d *ParallelDecoder) Register(obj interface{}) error {
	t := reflect.TypeOf(obj)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	d.structType = t

	// 读取表头
	if err := d.parseHeaders(); err != nil {
		return err
	}

	// 解析字段映射
	return d.parseStructMapping()
}

func (d *ParallelDecoder) parseHeaders() error {
	rows, err := d.file.Rows(d.sheet)
	if err != nil {
		return err
	}
	defer rows.Close()

	currentRow := 1
	for rows.Next() {
		if currentRow == d.headerRow {
			headers, _ := rows.Columns()
			for i, h := range headers {
				d.colMap[strings.TrimSpace(h)] = i
			}
			return nil
		}
		currentRow++
	}
	return fmt.Errorf("header row %d not found", d.headerRow)
}

func (d *ParallelDecoder) parseStructMapping() error {
	t := d.structType
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}

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
		d.fieldMetas = append(d.fieldMetas, &FieldMeta{
			ColName:    colName,
			Index:      colIdx,
			DateFormat: field.Tag.Get("format"),
			Default:    field.Tag.Get("default"),
			FieldIdx:   i,
			FieldType:  field.Type,
		})
	}
	return nil
}

type rowData struct {
	rowNum int
	cols   []string
}

// DecodeAllParallel 并行解码所有行
func (d *ParallelDecoder) DecodeAllParallel(result interface{}) error {
	resultPtr := reflect.ValueOf(result)
	if resultPtr.Kind() != reflect.Ptr || resultPtr.Elem().Kind() != reflect.Slice {
		return fmt.Errorf("result must be pointer to slice")
	}
	slicePtr := resultPtr.Elem()
	elemType := slicePtr.Type().Elem()

	// 读取所有行
	rows, err := d.file.Rows(d.sheet)
	if err != nil {
		return err
	}
	defer rows.Close()

	// 跳过表头
	for i := 1; i < d.headerRow; i++ {
		if !rows.Next() {
			return fmt.Errorf("not enough rows for header")
		}
	}
	rows.Next()

	allRows := make([]rowData, 0)
	rowNum := d.headerRow + 1
	for rows.Next() {
		cols, _ := rows.Columns()
		allRows = append(allRows, rowData{rowNum: rowNum, cols: cols})
		rowNum++
	}

	if len(allRows) == 0 {
		return nil
	}

	// 分批
	batches := d.splitIntoBatches(allRows)

	// 并行处理
	results := make(chan []interface{}, d.workers)
	var wg sync.WaitGroup

	for i := 0; i < d.workers && i < len(batches); i++ {
		wg.Add(1)
		go d.processBatch(batches[i], elemType, results, &wg)
	}

	go func() {
		wg.Wait()
		close(results)
	}()

	// 收集结果
	allResults := make([]interface{}, 0)
	for res := range results {
		allResults = append(allResults, res...)
	}

	// 按行号排序
	sort.Slice(allResults, func(i, j int) bool {
		// 假设实现了RowNumberer接口
		return false
	})

	// 写入结果
	for _, item := range allResults {
		slicePtr.Set(reflect.Append(slicePtr, reflect.ValueOf(item)))
	}

	return nil
}

func (d *ParallelDecoder) splitIntoBatches(rows []rowData) [][]rowData {
	batches := make([][]rowData, 0)
	for i := 0; i < len(rows); i += d.batchSize {
		end := i + d.batchSize
		if end > len(rows) {
			end = len(rows)
		}
		batches = append(batches, rows[i:end])
	}
	return batches
}

func (d *ParallelDecoder) processBatch(batch []rowData, elemType reflect.Type, results chan<- []interface{}, wg *sync.WaitGroup) {
	defer wg.Done()

	batchResults := make([]interface{}, 0, len(batch))

	for _, row := range batch {
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
			if meta.Index < len(row.cols) {
				cellValue = row.cols[meta.Index]
			}

			d.setFieldValueParallel(field, cellValue, meta)
		}

		if elemType.Kind() == reflect.Ptr {
			batchResults = append(batchResults, elemPtr.Interface())
		} else {
			batchResults = append(batchResults, elem.Interface())
		}
	}

	results <- batchResults
}

func (d *ParallelDecoder) setFieldValueParallel(field reflect.Value, cellValue string, meta *FieldMeta) {
	if cellValue == "" && meta.Default != "" {
		cellValue = meta.Default
	}
	if cellValue == "" {
		return
	}

	if meta.DateFormat != "" || meta.FieldType == reflect.TypeOf(time.Time{}) {
		if t, ok := TryParseExcelDate(cellValue, meta.DateFormat, false, 0); ok {
			field.Set(reflect.ValueOf(t))
		}
		return
	}

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
