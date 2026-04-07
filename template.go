package excel

import (
	"fmt"
	"html/template"
	"os"
)

// ========== 模板导出 ==========

// TemplateExporter 模板导出器
type TemplateExporter struct {
	templatePath string
	data         map[string]interface{}
	functions    template.FuncMap
}

// NewTemplateExporter 创建模板导出器
func NewTemplateExporter(templatePath string) *TemplateExporter {
	return &TemplateExporter{
		templatePath: templatePath,
		data:         make(map[string]interface{}),
		functions:    make(template.FuncMap),
	}
}

// SetData 设置数据
func (t *TemplateExporter) SetData(key string, value interface{}) *TemplateExporter {
	t.data[key] = value
	return t
}

// SetAllData 批量设置数据
func (t *TemplateExporter) SetAllData(data map[string]interface{}) *TemplateExporter {
	t.data = data
	return t
}

// AddFunc 添加自定义模板函数
func (t *TemplateExporter) AddFunc(name string, fn interface{}) *TemplateExporter {
	t.functions[name] = fn
	return t
}

// Export 导出到文件
func (t *TemplateExporter) Export(outputPath string) error {
	templateBytes, err := os.ReadFile(t.templatePath)
	if err != nil {
		return fmt.Errorf("读取模板文件失败: %w", err)
	}

	tmpl := template.New("excel_template").Funcs(t.functions)
	tmpl, err = tmpl.Parse(string(templateBytes))
	if err != nil {
		return fmt.Errorf("解析模板失败: %w", err)
	}

	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("创建输出文件失败: %w", err)
	}
	defer outFile.Close()

	if err := tmpl.Execute(outFile, t.data); err != nil {
		return fmt.Errorf("渲染模板失败: %w", err)
	}

	return nil
}
