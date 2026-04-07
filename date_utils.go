package excel

import (
	"time"
)

// ========== Excel日期处理（核心！） ==========

// Excel 日期存储为浮点数：
// - Windows Excel: 1900-01-01 = 1（有bug，1900-02-28 = 59, 1900-03-01 = 61）
// - Mac Excel 1904系统: 1904-01-01 = 0

const (
	excelEpochOffset = 25569 // 1970-01-01 距离 1900-01-01 的天数
	secondsPerDay    = 86400
)

// ExcelTimeToTime 将 Excel 的浮点数日期转换为 time.Time
// use1904: Mac Excel 使用1904日期系统时传 true
func ExcelTimeToTime(excelFloat float64, use1904 bool) time.Time {
	var days int64
	if use1904 {
		// 1904 日期系统
		days = int64(excelFloat)
		return time.Date(1904, 1, 1, 0, 0, 0, 0, time.Local).AddDate(0, 0, int(days))
	}

	// 1900 日期系统（Windows Excel 默认）
	days = int64(excelFloat)

	// 修正 Excel 的 1900-02-28 bug（Excel 认为 1900 年是闰年）
	if days > 59 {
		days -= 1 // 跳过不存在的 1900-02-29
	}

	// 1970-01-01 是 Excel 的 25569 天
	unixDays := days - excelEpochOffset
	return time.Unix(unixDays*secondsPerDay, 0).In(time.Local)
}

// TimeToExcelTime 将 time.Time 转换为 Excel 浮点数
func TimeToExcelTime(t time.Time, use1904 bool) float64 {
	if use1904 {
		days := t.Sub(time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)).Hours() / 24
		return days
	}

	days := t.Sub(time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC)).Hours()/24 + excelEpochOffset

	// 修正 Excel 的 1900-02-28 bug
	if days >= 61 {
		days += 1
	}
	return days
}

// TryParseExcelDate 尝试解析单元格值为时间（自动识别浮点数或字符串）
func TryParseExcelDate(cellValue string, dateFormat string, isExcelNumber bool, excelNum float64) (time.Time, bool) {
	// 优先处理 Excel 数字格式
	if isExcelNumber && excelNum >= 1 && excelNum <= 100000 {
		return ExcelTimeToTime(excelNum, false), true
	}

	// 字符串格式解析
	if dateFormat != "" {
		if t, err := time.Parse(dateFormat, cellValue); err == nil {
			return t, true
		}
	}

	// 常见日期格式
	commonFormats := []string{
		"2006-01-02",
		"2006/01/02",
		"2006年01月02日",
		"2006-01-02 15:04:05",
		"2006/01/02 15:04:05",
		"15:04:05",
		"20060102",
	}
	for _, f := range commonFormats {
		if t, err := time.Parse(f, cellValue); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}
