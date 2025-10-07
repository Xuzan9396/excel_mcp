package main

import (
    "encoding/json"
    "fmt"
    "log"
    "os"

    "github.com/Xuzan9396/excel_mcp/internal/excel"
)

func main() {
    log.Println("========== Excel MCP 功能测试 ==========")

    // 测试数据
    testData := []map[string]interface{}{
        {"姓名": "张三", "年龄": 25, "城市": "北京"},
        {"姓名": "李四", "年龄": 30, "城市": "上海"},
        {"姓名": "王五", "年龄": 28, "城市": "广州"},
    }

    // 测试 1: 写入和读取 XLSX
    log.Println("\n[测试 1] 写入和读取 XLSX 文件")
    xlsxFile := "test_output.xlsx"
    if err := testXLSX(xlsxFile, testData); err != nil {
        log.Printf("❌ XLSX 测试失败: %v", err)
    } else {
        log.Println("✅ XLSX 测试成功")
    }

    // 测试 2: 写入和读取 CSV
    log.Println("\n[测试 2] 写入和读取 CSV 文件")
    csvFile := "test_output.csv"
    if err := testCSV(csvFile, testData); err != nil {
        log.Printf("❌ CSV 测试失败: %v", err)
    } else {
        log.Println("✅ CSV 测试成功")
    }

    // 测试 3: XLSX 多 Sheet
    log.Println("\n[测试 3] XLSX 多 Sheet 测试")
    multiSheetFile := "test_multi_sheet.xlsx"
    if err := testMultiSheet(multiSheetFile, testData); err != nil {
        log.Printf("❌ 多 Sheet 测试失败: %v", err)
    } else {
        log.Println("✅ 多 Sheet 测试成功")
    }

    log.Println("\n========== 测试完成 ==========")
}

func testXLSX(filePath string, testData []map[string]interface{}) error {
    // 写入
    log.Printf("写入 XLSX: %s", filePath)
    if err := excel.WriteExcel(filePath, testData, "测试数据"); err != nil {
        return fmt.Errorf("写入失败: %w", err)
    }

    // 读取
    log.Printf("读取 XLSX: %s", filePath)
    data, err := excel.ReadExcel(filePath, "测试数据")
    if err != nil {
        return fmt.Errorf("读取失败: %w", err)
    }

    // 验证
    jsonData, _ := json.MarshalIndent(data, "", "  ")
    log.Printf("读取的数据:\n%s", string(jsonData))

    if len(data) != len(testData) {
        return fmt.Errorf("数据行数不匹配: 期望 %d, 实际 %d", len(testData), len(data))
    }

    // 清理
    os.Remove(filePath)

    return nil
}

func testCSV(filePath string, testData []map[string]interface{}) error {
    // 写入
    log.Printf("写入 CSV: %s", filePath)
    if err := excel.WriteExcel(filePath, testData, ""); err != nil {
        return fmt.Errorf("写入失败: %w", err)
    }

    // 读取
    log.Printf("读取 CSV: %s", filePath)
    data, err := excel.ReadExcel(filePath, "")
    if err != nil {
        return fmt.Errorf("读取失败: %w", err)
    }

    // 验证
    jsonData, _ := json.MarshalIndent(data, "", "  ")
    log.Printf("读取的数据:\n%s", string(jsonData))

    if len(data) != len(testData) {
        return fmt.Errorf("数据行数不匹配: 期望 %d, 实际 %d", len(testData), len(data))
    }

    // 清理
    os.Remove(filePath)

    return nil
}

func testMultiSheet(filePath string, testData []map[string]interface{}) error {
    // 写入第一个 sheet
    log.Printf("写入第一个 Sheet")
    if err := excel.WriteExcel(filePath, testData, "Sheet1"); err != nil {
        return fmt.Errorf("写入 Sheet1 失败: %w", err)
    }

    // 读取第一个 sheet（不指定名称，应该读取第一个）
    log.Printf("读取第一个 Sheet（默认）")
    data1, err := excel.ReadExcel(filePath, "")
    if err != nil {
        return fmt.Errorf("读取失败: %w", err)
    }

    jsonData1, _ := json.MarshalIndent(data1, "", "  ")
    log.Printf("Sheet1 数据:\n%s", string(jsonData1))

    // 读取指定 sheet
    log.Printf("读取指定 Sheet（Sheet1）")
    data2, err := excel.ReadExcel(filePath, "Sheet1")
    if err != nil {
        return fmt.Errorf("读取 Sheet1 失败: %w", err)
    }

    if len(data2) != len(testData) {
        return fmt.Errorf("数据行数不匹配: 期望 %d, 实际 %d", len(testData), len(data2))
    }

    // 清理
    os.Remove(filePath)

    return nil
}

