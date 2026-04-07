package excel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ========== 数据校验 ==========

// ValidationRule 校验规则
type ValidationRule struct {
	Required bool     // 必填
	Min      *float64 // 最小值
	Max      *float64 // 最大值
	MinLen   *int     // 最小长度
	MaxLen   *int     // 最大长度
	ExactLen *int     // 固定长度
	Regex    string   // 正则表达式
	InValues []string // 枚举值列表
	Email    bool     // 邮箱格式
	Phone    bool     // 手机号格式
	IDCard   bool     // 身份证格式
}

// Validator 校验器
type Validator struct {
	collector *ErrorCollector
}

// NewValidator 创建校验器
func NewValidator() *Validator {
	return &Validator{
		collector: NewErrorCollector(),
	}
}

// ParseValidationTag 解析校验标签
func ParseValidationTag(tag string) *ValidationRule {
	if tag == "" {
		return nil
	}

	rule := &ValidationRule{}
	parts := strings.Split(tag, ";")

	for _, part := range parts {
		part = strings.TrimSpace(part)

		switch {
		case part == "required":
			rule.Required = true

		case strings.HasPrefix(part, "min="):
			val, err := strconv.ParseFloat(strings.TrimPrefix(part, "min="), 64)
			if err == nil {
				rule.Min = &val
			}

		case strings.HasPrefix(part, "max="):
			val, err := strconv.ParseFloat(strings.TrimPrefix(part, "max="), 64)
			if err == nil {
				rule.Max = &val
			}

		case strings.HasPrefix(part, "len="):
			val, err := strconv.Atoi(strings.TrimPrefix(part, "len="))
			if err == nil {
				rule.ExactLen = &val
			}

		case strings.HasPrefix(part, "minlen="):
			val, err := strconv.Atoi(strings.TrimPrefix(part, "minlen="))
			if err == nil {
				rule.MinLen = &val
			}

		case strings.HasPrefix(part, "maxlen="):
			val, err := strconv.Atoi(strings.TrimPrefix(part, "maxlen="))
			if err == nil {
				rule.MaxLen = &val
			}

		case strings.HasPrefix(part, "regex="):
			rule.Regex = strings.TrimPrefix(part, "regex=")

		case strings.HasPrefix(part, "in="):
			values := strings.TrimPrefix(part, "in=")
			rule.InValues = strings.Split(values, ",")

		case part == "email":
			rule.Email = true

		case part == "phone":
			rule.Phone = true

		case part == "idcard":
			rule.IDCard = true
		}
	}

	return rule
}

// Validate 执行校验，返回错误信息列表
func (v *Validator) Validate(value string, rule *ValidationRule) []string {
	var errors []string

	if rule == nil {
		return errors
	}

	// 必填校验
	if rule.Required && value == "" {
		errors = append(errors, "该字段为必填项")
		return errors
	}

	// 空值跳过后续校验
	if value == "" {
		return errors
	}

	// 固定长度校验
	if rule.ExactLen != nil && len(value) != *rule.ExactLen {
		errors = append(errors, fmt.Sprintf("长度必须为 %d 个字符", *rule.ExactLen))
	}

	// 最小长度校验
	if rule.MinLen != nil && len(value) < *rule.MinLen {
		errors = append(errors, fmt.Sprintf("长度不能少于 %d 个字符", *rule.MinLen))
	}

	// 最大长度校验
	if rule.MaxLen != nil && len(value) > *rule.MaxLen {
		errors = append(errors, fmt.Sprintf("长度不能超过 %d 个字符", *rule.MaxLen))
	}

	// 数值范围校验
	if rule.Min != nil || rule.Max != nil {
		if num, err := strconv.ParseFloat(value, 64); err == nil {
			if rule.Min != nil && num < *rule.Min {
				errors = append(errors, fmt.Sprintf("数值不能小于 %.2f", *rule.Min))
			}
			if rule.Max != nil && num > *rule.Max {
				errors = append(errors, fmt.Sprintf("数值不能大于 %.2f", *rule.Max))
			}
		}
	}

	// 正则校验
	if rule.Regex != "" {
		matched, err := regexp.MatchString(rule.Regex, value)
		if err == nil && !matched {
			errors = append(errors, fmt.Sprintf("格式不符合要求: %s", rule.Regex))
		}
	}

	// 枚举校验
	if len(rule.InValues) > 0 {
		found := false
		for _, v := range rule.InValues {
			if v == value {
				found = true
				break
			}
		}
		if !found {
			errors = append(errors, fmt.Sprintf("值必须为: %s", strings.Join(rule.InValues, ", ")))
		}
	}

	// 邮箱校验
	if rule.Email {
		emailRegex := `^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$`
		matched, _ := regexp.MatchString(emailRegex, value)
		if !matched {
			errors = append(errors, "邮箱格式不正确")
		}
	}

	// 手机号校验
	if rule.Phone {
		phoneRegex := `^1[3-9]\d{9}$`
		matched, _ := regexp.MatchString(phoneRegex, value)
		if !matched {
			errors = append(errors, "手机号格式不正确（11位数字，以1开头）")
		}
	}

	// 身份证校验
	if rule.IDCard {
		idRegex := `^[1-9]\d{5}(18|19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{3}[\dXx]$`
		matched, _ := regexp.MatchString(idRegex, value)
		if !matched {
			errors = append(errors, "身份证号格式不正确")
		}
	}

	return errors
}
