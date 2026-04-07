package excel

import (
	"fmt"
	"strings"
	"sync"
)

// ========== 错误收集器 ==========

type ErrorLevel string

const (
	ErrorLevelInfo    ErrorLevel = "INFO"
	ErrorLevelWarning ErrorLevel = "WARNING"
	ErrorLevelError   ErrorLevel = "ERROR"
	ErrorLevelFatal   ErrorLevel = "FATAL"
)

// ExcelError Excel处理错误
type ExcelError struct {
	Level    ErrorLevel `json:"level"`
	Row      int        `json:"row"`
	Column   string     `json:"column"`
	Field    string     `json:"field"`
	Message  string     `json:"message"`
	RawValue string     `json:"raw_value"`
}

func (e *ExcelError) Error() string {
	return fmt.Sprintf("[%s] row=%d, col=%s, field=%s, value=%q: %s",
		e.Level, e.Row, e.Column, e.Field, e.RawValue, e.Message)
}

// ErrorCollector 错误收集器
type ErrorCollector struct {
	mu          sync.RWMutex
	errors      []*ExcelError
	StopOnError bool // 遇到ERROR级别是否停止
	MaxErrors   int  // 最大错误数（0表示无限制）
}

// NewErrorCollector 创建错误收集器
func NewErrorCollector() *ErrorCollector {
	return &ErrorCollector{
		errors:      make([]*ExcelError, 0),
		StopOnError: false,
		MaxErrors:   100,
	}
}

// Add 添加错误
func (c *ErrorCollector) Add(err *ExcelError) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.errors = append(c.errors, err)
}

// HasError 是否有错误
func (c *ErrorCollector) HasError() bool {
	c.mu.RLock()
	defer c.mu.RUnlock()
	for _, e := range c.errors {
		if e.Level == ErrorLevelError || e.Level == ErrorLevelFatal {
			return true
		}
	}
	return false
}

// HasFatal 是否有致命错误
func (c *ErrorCollector) HasFatal() bool {
	c.mu.RLock()
	defer c.mu.RUnlock()
	for _, e := range c.errors {
		if e.Level == ErrorLevelFatal {
			return true
		}
	}
	return false
}

// ShouldStop 是否应该停止
func (c *ErrorCollector) ShouldStop() bool {
	if !c.StopOnError {
		return false
	}
	c.mu.RLock()
	defer c.mu.RUnlock()
	for _, e := range c.errors {
		if e.Level == ErrorLevelError || e.Level == ErrorLevelFatal {
			return true
		}
	}
	return false
}

// ShouldContinue 是否应该继续
func (c *ErrorCollector) ShouldContinue() bool {
	if c.MaxErrors > 0 && len(c.errors) >= c.MaxErrors {
		return false
	}
	return !c.ShouldStop()
}

// Errors 获取所有错误
func (c *ErrorCollector) Errors() []*ExcelError {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return append([]*ExcelError{}, c.errors...)
}

// Summary 错误摘要
func (c *ErrorCollector) Summary() string {
	c.mu.RLock()
	defer c.mu.RUnlock()

	counts := make(map[ErrorLevel]int)
	for _, e := range c.errors {
		counts[e.Level]++
	}

	var parts []string
	for level, count := range counts {
		parts = append(parts, fmt.Sprintf("%s:%d", level, count))
	}
	return strings.Join(parts, ", ")
}

// Clear 清空错误
func (c *ErrorCollector) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.errors = make([]*ExcelError, 0)
}

// Count 错误数量
func (c *ErrorCollector) Count() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.errors)
}
