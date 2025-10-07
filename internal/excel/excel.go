package excel

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ReadExcel 读取 Excel 文件（支持 xlsx 和 csv）
// filePath: 文件路径
// sheetName: sheet 名称，对于 xlsx 文件，如果为空则读取第一个 sheet；对于 csv 文件，此参数无效
// 返回格式: []map[string]interface{}
func ReadExcel(filePath string, sheetName string) ([]map[string]interface{}, error) {
	ext := strings.ToLower(filepath.Ext(filePath))

	switch ext {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return readXLSX(filePath, sheetName)
	case ".csv":
		return readCSV(filePath)
	default:
		return nil, fmt.Errorf("unsupported file format: %s", ext)
	}
}

// WriteExcel 写入 Excel 文件（支持 xlsx 和 csv）
// filePath: 文件路径
// data: 数据，格式为 []map[string]interface{}，其中 key 为列标题
// sheetName: sheet 名称，对于 xlsx 文件，如果为空则使用 "Sheet1"；对于 csv 文件，此参数无效
func WriteExcel(filePath string, data []map[string]interface{}, sheetName string) error {
	if len(data) == 0 {
		return fmt.Errorf("no data to write")
	}

	ext := strings.ToLower(filepath.Ext(filePath))

	switch ext {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return writeXLSX(filePath, data, sheetName)
	case ".csv":
		return writeCSV(filePath, data)
	default:
		return fmt.Errorf("unsupported file format: %s", ext)
	}
}

// readXLSX 读取 xlsx 文件
func readXLSX(filePath string, sheetName string) ([]map[string]interface{}, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("warning: failed to close file: %v\n", err)
		}
	}()

	// 如果没有指定 sheet 名称，获取第一个 sheet
	if sheetName == "" {
		sheetName = f.GetSheetName(0)
		if sheetName == "" {
			return nil, fmt.Errorf("no sheets found in file")
		}
	}

	// 读取所有行
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to read rows: %w", err)
	}

	if len(rows) == 0 {
		return nil, fmt.Errorf("no data found in sheet: %s", sheetName)
	}

	// 第一行作为标题
	headers := rows[0]
	if len(headers) == 0 {
		return nil, fmt.Errorf("no headers found in sheet: %s", sheetName)
	}

	// 转换为 map 格式
	var result []map[string]interface{}
	for i := 1; i < len(rows); i++ {
		row := rows[i]
		rowData := make(map[string]interface{})

		for j, header := range headers {
			if header == "" {
				continue // 跳过空标题列
			}

			var value interface{}
			if j < len(row) {
				value = parseValue(row[j])
			} else {
				value = ""
			}
			rowData[header] = value
		}

		result = append(result, rowData)
	}

	return result, nil
}

// readCSV 读取 csv 文件
func readCSV(filePath string) ([]map[string]interface{}, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return nil, fmt.Errorf("failed to read CSV: %w", err)
	}

	if len(records) == 0 {
		return nil, fmt.Errorf("no data found in CSV file")
	}

	// 第一行作为标题
	headers := records[0]
	if len(headers) == 0 {
		return nil, fmt.Errorf("no headers found in CSV file")
	}

	// 转换为 map 格式
	var result []map[string]interface{}
	for i := 1; i < len(records); i++ {
		row := records[i]
		rowData := make(map[string]interface{})

		for j, header := range headers {
			if header == "" {
				continue
			}

			var value interface{}
			if j < len(row) {
				value = parseValue(row[j])
			} else {
				value = ""
			}
			rowData[header] = value
		}

		result = append(result, rowData)
	}

	return result, nil
}

// writeXLSX 写入 xlsx 文件
func writeXLSX(filePath string, data []map[string]interface{}, sheetName string) error {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	// 创建新文件
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("warning: failed to close file: %v\n", err)
		}
	}()

	// 创建或重命名 sheet
	defaultSheet := f.GetSheetName(0)
	if defaultSheet != "" && defaultSheet != sheetName {
		f.SetSheetName(defaultSheet, sheetName)
	}

	// 获取所有列名（从第一行数据中提取）
	var headers []string
	headerMap := make(map[string]bool)
	for _, row := range data {
		for key := range row {
			if !headerMap[key] {
				headers = append(headers, key)
				headerMap[key] = true
			}
		}
	}

	// 写入标题行
	for i, header := range headers {
		cell, err := excelize.CoordinatesToCellName(i+1, 1)
		if err != nil {
			return fmt.Errorf("failed to get cell name: %w", err)
		}
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return fmt.Errorf("failed to set header: %w", err)
		}
	}

	// 写入数据行
	for rowIdx, rowData := range data {
		for colIdx, header := range headers {
			cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			if err != nil {
				return fmt.Errorf("failed to get cell name: %w", err)
			}

			value := rowData[header]
			if value == nil {
				value = ""
			}

			// 设置单元格值和格式，防止大数字转科学计数法
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}

			// 为数字类型设置格式，防止科学计数法
			switch v := value.(type) {
			case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
				// 整数：设置为数字格式，不使用科学计数法
				style, err := f.NewStyle(&excelize.Style{
					NumFmt: 1, // 0.00 格式，但对于整数会显示为整数
				})
				if err == nil {
					f.SetCellStyle(sheetName, cell, cell, style)
				}
			case float32:
				// 浮点数：检查是否是整数值
				if v == float32(int64(v)) {
					// 如果是整数值，设置为整数格式
					style, err := f.NewStyle(&excelize.Style{
						NumFmt: 1,
					})
					if err == nil {
						f.SetCellStyle(sheetName, cell, cell, style)
					}
				}
			case float64:
				// 浮点数：检查是否是整数值
				if v == float64(int64(v)) {
					// 如果是整数值，设置为整数格式
					style, err := f.NewStyle(&excelize.Style{
						NumFmt: 1,
					})
					if err == nil {
						f.SetCellStyle(sheetName, cell, cell, style)
					}
				}
			}
		}
	}

	// 保存文件
	if err := f.SaveAs(filePath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	return nil
}

// writeCSV 写入 csv 文件
func writeCSV(filePath string, data []map[string]interface{}) error {
	file, err := os.Create(filePath)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	// 获取所有列名
	var headers []string
	headerMap := make(map[string]bool)
	for _, row := range data {
		for key := range row {
			if !headerMap[key] {
				headers = append(headers, key)
				headerMap[key] = true
			}
		}
	}

	// 写入标题行
	if err := writer.Write(headers); err != nil {
		return fmt.Errorf("failed to write headers: %w", err)
	}

	// 写入数据行
	for _, rowData := range data {
		var row []string
		for _, header := range headers {
			value := rowData[header]
			row = append(row, formatValue(value))
		}
		if err := writer.Write(row); err != nil {
			return fmt.Errorf("failed to write row: %w", err)
		}
	}

	return nil
}

// formatValue 格式化值，防止科学计数法
func formatValue(value interface{}) string {
	if value == nil {
		return ""
	}

	switch v := value.(type) {
	case int:
		return fmt.Sprintf("%d", v)
	case int8:
		return fmt.Sprintf("%d", v)
	case int16:
		return fmt.Sprintf("%d", v)
	case int32:
		return fmt.Sprintf("%d", v)
	case int64:
		return fmt.Sprintf("%d", v)
	case uint:
		return fmt.Sprintf("%d", v)
	case uint8:
		return fmt.Sprintf("%d", v)
	case uint16:
		return fmt.Sprintf("%d", v)
	case uint32:
		return fmt.Sprintf("%d", v)
	case uint64:
		return fmt.Sprintf("%d", v)
	case float32:
		// 检查是否是整数值
		if v == float32(int64(v)) {
			return fmt.Sprintf("%.0f", v)
		}
		return fmt.Sprintf("%f", v)
	case float64:
		// 检查是否是整数值
		if v == float64(int64(v)) {
			return fmt.Sprintf("%.0f", v)
		}
		return fmt.Sprintf("%f", v)
	case string:
		return v
	case bool:
		if v {
			return "true"
		}
		return "false"
	default:
		return fmt.Sprintf("%v", v)
	}
}

// parseValue 尝试解析值的类型
func parseValue(s string) interface{} {
	if s == "" {
		return ""
	}

	// 尝试解析为数字
	var i int
	if _, err := fmt.Sscanf(s, "%d", &i); err == nil {
		return i
	}

	var f float64
	if _, err := fmt.Sscanf(s, "%f", &f); err == nil {
		return f
	}

	// 返回字符串
	return s
}
