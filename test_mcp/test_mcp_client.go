package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"

	"github.com/Xuzan9396/excel_mcp/internal/excel"
	"github.com/mark3labs/mcp-go/client"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

func main() {
	log.Println("========== MCP Server 测试 ==========")

	// 创建测试数据
	testData := []map[string]interface{}{
		{"产品": "iPhone", "价格": 6999, "库存": 100},
		{"产品": "iPad", "价格": 3999, "库存": 50},
		{"产品": "MacBook", "价格": 12999, "库存": 30},
	}

	testFile := "test_mcp_data.xlsx"
	defer os.Remove(testFile)

	// 先写入测试文件
	log.Println("\n[准备] 创建测试文件")
	if err := excel.WriteExcel(testFile, testData, "产品数据"); err != nil {
		log.Fatalf("创建测试文件失败: %v", err)
	}
	log.Printf("✅ 测试文件创建成功: %s", testFile)

	// 创建 MCP Server
	s := createMCPServer()

	// 创建 In-Process Client
	mcpClient, err := client.NewInProcessClient(s)
	if err != nil {
		log.Fatalf("创建客户端失败: %v", err)
	}
	defer mcpClient.Close()

	ctx := context.Background()

	// 初始化客户端
	log.Println("\n[步骤 1] 初始化 MCP 客户端")
	initReq := mcp.InitializeRequest{}
	initReq.Params.ProtocolVersion = "2024-11-05"
	initReq.Params.ClientInfo.Name = "test-client"
	initReq.Params.ClientInfo.Version = "1.0.0"

	_, err = mcpClient.Initialize(ctx, initReq)
	if err != nil {
		log.Fatalf("初始化失败: %v", err)
	}
	log.Println("✅ 初始化成功")

	// 测试 1: 读取 Excel
	log.Println("\n[测试 1] 调用 read_excel 工具")
	req1 := mcp.CallToolRequest{}
	req1.Params.Name = "read_excel"
	req1.Params.Arguments = map[string]interface{}{
		"file_path":  testFile,
		"sheet_name": "产品数据",
	}

	result, err := mcpClient.CallTool(ctx, req1)
	if err != nil {
		log.Fatalf("调用失败: %v", err)
	}

	if len(result.Content) > 0 {
		if textContent, ok := mcp.AsTextContent(result.Content[0]); ok {
			log.Printf("读取结果:\n%s", textContent.Text)

			// 验证数据
			var readData []map[string]interface{}
			if err := json.Unmarshal([]byte(textContent.Text), &readData); err != nil {
				log.Fatalf("解析数据失败: %v", err)
			}

			if len(readData) != len(testData) {
				log.Fatalf("❌ 数据行数不匹配: 期望 %d, 实际 %d", len(testData), len(readData))
			}
			log.Println("✅ 数据验证成功")
		}
	}

	// 测试 2: 写入 Excel
	log.Println("\n[测试 2] 调用 write_excel 工具")
	writeFile := "test_mcp_write.xlsx"
	defer os.Remove(writeFile)

	writeData := []map[string]interface{}{
		{"员工": "张三", "部门": "技术部", "工资": 15000},
		{"员工": "李四", "部门": "市场部", "工资": 12000},
	}

	jsonData, _ := json.Marshal(writeData)

	req2 := mcp.CallToolRequest{}
	req2.Params.Name = "write_excel"
	req2.Params.Arguments = map[string]interface{}{
		"file_path":  writeFile,
		"data":       string(jsonData),
		"sheet_name": "员工信息",
	}

	result2, err := mcpClient.CallTool(ctx, req2)
	if err != nil {
		log.Fatalf("调用失败: %v", err)
	}

	if len(result2.Content) > 0 {
		if textContent, ok := mcp.AsTextContent(result2.Content[0]); ok {
			log.Printf("写入结果: %s", textContent.Text)
		}
	}

	// 验证写入的文件
	log.Println("\n[验证] 读取刚写入的文件")
	verifyData, err := excel.ReadExcel(writeFile, "员工信息")
	if err != nil {
		log.Fatalf("读取失败: %v", err)
	}

	verifyJSON, _ := json.MarshalIndent(verifyData, "", "  ")
	log.Printf("验证数据:\n%s", string(verifyJSON))

	if len(verifyData) != len(writeData) {
		log.Fatalf("❌ 数据行数不匹配")
	}
	log.Println("✅ 写入验证成功")

	log.Println("\n========== 所有测试通过 ==========")
}

func createMCPServer() *server.MCPServer {
	s := server.NewMCPServer(
		"Excel MCP Server",
		"1.0.0",
		server.WithToolCapabilities(true),
	)

	// 注册 read_excel 工具
	s.AddTool(
		mcp.NewTool("read_excel",
			mcp.WithDescription("读取 Excel 或 CSV 文件"),
			mcp.WithString("file_path", mcp.Required()),
			mcp.WithString("sheet_name"),
		),
		handleReadExcel,
	)

	// 注册 write_excel 工具
	s.AddTool(
		mcp.NewTool("write_excel",
			mcp.WithDescription("写入 Excel 或 CSV 文件"),
			mcp.WithString("file_path", mcp.Required()),
			mcp.WithString("data", mcp.Required()),
			mcp.WithString("sheet_name"),
		),
		handleWriteExcel,
	)

	return s
}

func handleReadExcel(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, _ := request.RequireString("file_path")
	sheetName := request.GetString("sheet_name", "")

	data, err := excel.ReadExcel(filePath, sheetName)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("读取失败: %v", err)), nil
	}

	jsonData, _ := json.MarshalIndent(data, "", "  ")
	return mcp.NewToolResultText(string(jsonData)), nil
}

func handleWriteExcel(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, _ := request.RequireString("file_path")
	dataStr, _ := request.RequireString("data")
	sheetName := request.GetString("sheet_name", "")

	var data []map[string]interface{}
	if err := json.Unmarshal([]byte(dataStr), &data); err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("JSON 解析失败: %v", err)), nil
	}

	if err := excel.WriteExcel(filePath, data, sheetName); err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("写入失败: %v", err)), nil
	}

	return mcp.NewToolResultText(fmt.Sprintf("✅ 文件写入成功: %s", filePath)), nil
}
