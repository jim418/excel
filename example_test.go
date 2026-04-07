package excel

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// Employee 员工结构体（带完整标签）
type Employee struct {
	ID        int       `excel:"工号" validate:"required;min=1000;max=99999"`
	Name      string    `excel:"姓名" validate:"required;minlen=2;maxlen=20"`
	Dept      string    `excel:"部门" validate:"in=技术部,市场部,销售部,人事部"`
	Salary    float64   `excel:"工资" validate:"required;min=3000;max=100000"`
	Email     string    `excel:"邮箱" validate:"email"`
	Phone     string    `excel:"手机号" validate:"phone"`
	HireDate  time.Time `excel:"入职日期" format:"2006-01-02" validate:"required"`
	Status    string    `excel:"状态" default:"试用期" validate:"in=试用期,正式,离职"`
}

// ExampleBasicEncoder 基础编码器示例
func ExampleBasicEncoder() {
	fmt.Println("=== 基础编码器示例 ===")

	employees := []Employee{
		{ID: 1001, Name: "张三", Dept: "技术部", Salary: 15000,
			Email: "zhangsan@example.com", Phone: "13812345678",
			HireDate: time.Date(2023, 1, 15, 0, 0, 0, 0, time.Local), Status: "正式"},
		{ID: 1002, Name: "李四", Dept: "市场部", Salary: 12000,
			Email: "lisi@example.com", Phone: "13912345678",
			HireDate: time.Date(2023, 3, 20, 0, 0, 0, 0, time.Local), Status: "试用期"},
	}

	f := excelize.NewFile()
	defer f.Close()

	sheet := "员工表"
	f.NewSheet(sheet)
	f.DeleteSheet("Sheet1")

	encoder := NewEncoder(f, sheet)
	encoder.SetStartRow(1)

	if err := encoder.Register(Employee{}); err != nil {
		panic(err)
	}
	if err := encoder.WriteHeaders(); err != nil {
		panic(err)
	}
	if err := encoder.EncodeAll(employees); err != nil {
		panic(err)
	}

	if err := f.SaveAs("basic_export.xlsx"); err != nil {
		panic(err)
	}
	fmt.Println("导出成功: basic_export.xlsx")
}

// ExampleStreamEncoder 流式编码器示例（大数据量）
func ExampleStreamEncoder() {
	fmt.Println("\n=== 流式编码器示例 ===")

	// 生成10万条数据
	employees := make([]Employee, 100000)
	for i := 0; i < 100000; i++ {
		employees[i] = Employee{
			ID:       1000 + i,
			Name:     fmt.Sprintf("员工_%d", i+1),
			Dept:     []string{"技术部", "市场部", "销售部", "人事部"}[i%4],
			Salary:   float64(5000 + i%50000),
			Email:    fmt.Sprintf("user%d@example.com", i+1),
			Phone:    "13812345678",
			HireDate: time.Now().AddDate(0, 0, -i%365),
			Status:   "正式",
		}
	}

	f := excelize.NewFile()
	defer f.Close()

	sheet := "员工表"
	f.NewSheet(sheet)
	f.DeleteSheet("Sheet1")

	encoder, err := NewStreamEncoder(f, sheet)
	if err != nil {
		panic(err)
	}
	encoder.SetStartRow(1)

	if err := encoder.Register(Employee{}); err != nil {
		panic(err)
	}
	if err := encoder.WriteHeaders(); err != nil {
		panic(err)
	}
	if err := encoder.EncodeAll(employees); err != nil {
		panic(err)
	}
	if err := encoder.Flush(); err != nil {
		panic(err)
	}

	if err := f.SaveAs("stream_export.xlsx"); err != nil {
		panic(err)
	}
	fmt.Println("流式导出成功: stream_export.xlsx (10万条)")
}

// ExampleDecoder 基础解码器示例
func ExampleDecoder() {
	fmt.Println("\n=== 基础解码器示例 ===")

	f, err := excelize.OpenFile("basic_export.xlsx")
	if err != nil {
		fmt.Printf("打开文件失败: %v\n", err)
		return
	}
	defer f.Close()

	decoder := NewDecoder(f, "员工表")
	decoder.SetHeaderRow(1)

	if err := decoder.Register(Employee{}); err != nil {
		panic(err)
	}

	var employees []Employee
	if err := decoder.DecodeAll(&employees); err != nil {
		panic(err)
	}

	fmt.Printf("成功导入 %d 条数据\n", len(employees))
	for i, emp := range employees[:3] {
		fmt.Printf("  %d: ID=%d, 姓名=%s, 部门=%s, 工资=%.2f\n",
			i+1, emp.ID, emp.Name, emp.Dept, emp.Salary)
	}
}

// ExampleExcelValidation Excel原生数据验证示例
func ExampleExcelValidation() {
	fmt.Println("\n=== Excel原生数据验证示例 ===")

	f := excelize.NewFile()
	sheet := "数据录入"
	f.NewSheet(sheet)
	f.DeleteSheet("Sheet1")

	// 设置表头
	headers := []string{"姓名", "年龄", "部门", "工资", "入职日期"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, h)
	}

	// 部门下拉列表
	AddDropdownList(f, sheet, "C", 2, 100, []string{"技术部", "市场部", "销售部", "人事部"}, false)

	// 年龄范围验证
	AddRangeValidation(f, sheet, "B", 2, 100, 18, 60, true)

	// 工资范围验证
	AddRangeValidation(f, sheet, "D", 2, 100, 3000, 50000, false)

	// 入职日期范围验证
	AddDateRangeValidation(f, sheet, "E", 2, 100, "2020-01-01", "2025-12-31", true)

	if err := f.SaveAs("validated_template.xlsx"); err != nil {
		panic(err)
	}
	fmt.Println("已创建带数据验证的模板: validated_template.xlsx")
}

// ExampleErrorCollector 错误收集器示例
func ExampleErrorCollector() {
	fmt.Println("\n=== 错误收集器示例 ===")

	collector := NewErrorCollector()
	collector.StopOnError = false

	collector.Add(&ExcelError{
		Level:    ErrorLevelError,
		Row:      2,
		Column:   "姓名",
		Field:    "Name",
		RawValue: "",
		Message:  "姓名为必填项",
	})

	collector.Add(&ExcelError{
		Level:    ErrorLevelError,
		Row:      3,
		Column:   "年龄",
		Field:    "Age",
		RawValue: "150",
		Message:  "年龄超出范围",
	})

	fmt.Printf("错误数量: %d\n", collector.Count())
	fmt.Printf("错误摘要: %s\n", collector.Summary())
	fmt.Printf("是否有错误: %v\n", collector.HasError())

	for _, err := range collector.Errors() {
		fmt.Printf("  - %s\n", err.Error())
	}
}

// ExampleValidation 数据校验示例
func ExampleValidation() {
	fmt.Println("\n=== 数据校验示例 ===")

	validator := NewValidator()

	testCases := []struct {
		value string
		rule  *ValidationRule
	}{
		{"", &ValidationRule{Required: true}},
		{"abc", &ValidationRule{MinLen: intPtr(5), MaxLen: intPtr(10)}},
		{"150", &ValidationRule{Min: float64Ptr(0), Max: float64Ptr(100)}},
		{"invalid-email", &ValidationRule{Email: true}},
		{"13812345678", &ValidationRule{Phone: true}},
		{"技术部", &ValidationRule{InValues: []string{"技术部", "市场部"}}},
	}

	for i, tc := range testCases {
		errors := validator.Validate(tc.value, tc.rule)
		fmt.Printf("用例%d: 值=%q, 错误数=%d\n", i+1, tc.value, len(errors))
		for _, e := range errors {
			fmt.Printf("    - %s\n", e)
		}
	}
}

func intPtr(v int) *int         { return &v }
func float64Ptr(v float64) *float64 { return &v }

// RunAllExamples 运行所有示例
func RunAllExamples() {
	ExampleBasicEncoder()
	ExampleDecoder()
	ExampleStreamEncoder()
	ExampleExcelValidation()
	ExampleErrorCollector()
	ExampleValidation()
}
