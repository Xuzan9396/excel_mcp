package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"

	"github.com/Xuzan9396/excel_mcp/internal/excel"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
)

func main() {
	// 创建 MCP Server
	mcpServer := server.NewMCPServer(
		"Excel MCP Server",
		"1.0.0",
		server.WithToolCapabilities(true),
	)

	// 注册工具
	registerTools(mcpServer)

	// 启动 STDIO Server
	log.Println("Excel MCP Server 启动中...")
	if err := server.ServeStdio(mcpServer); err != nil {
		fmt.Fprintf(os.Stderr, "Server error: %v\n", err)
		os.Exit(1)
	}
}

// registerTools 注册所有工具
func registerTools(s *server.MCPServer) {
	// 1. read_excel 工具
	s.AddTool(
		mcp.NewTool("read_excel",
			mcp.WithDescription("读取 Excel 或 CSV 文件，返回 JSON 格式数据"),
			mcp.WithString("file_path",
				mcp.Required(),
				mcp.Description("文件路径（支持 .xlsx, .xlsm, .xltx, .xltm, .csv 格式）"),
			),
			mcp.WithString("sheet_name",
				mcp.Description("Sheet 名称（仅对 xlsx 文件有效，为空则读取第一个 sheet）"),
			),
		),
		handleReadExcel,
	)

	// 2. write_excel 工具
	s.AddTool(
		mcp.NewTool("write_excel",
			mcp.WithDescription("写入数据到 Excel 或 CSV 文件"),
			mcp.WithString("file_path",
				mcp.Required(),
				mcp.Description("文件路径（支持 .xlsx, .xlsm, .xltx, .xltm, .csv 格式）"),
			),
			mcp.WithString("data",
				mcp.Required(),
				mcp.Description("JSON 格式的数据，格式为 [{\"列名1\":值1,\"列名2\":值2}]"),
			),
			mcp.WithString("sheet_name",
				mcp.Description("Sheet 名称（仅对 xlsx 文件有效，为空则使用 Sheet1）"),
			),
		),
		handleWriteExcel,
	)
}

// handleReadExcel 处理读取 Excel
func handleReadExcel(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, err := request.RequireString("file_path")
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("file_path 参数错误: %v", err)), nil
	}

	sheetName := request.GetString("sheet_name", "")

	log.Printf("read_excel 工具被调用: file_path=%s, sheet_name=%s", filePath, sheetName)

	// 读取 Excel
	data, err := excel.ReadExcel(filePath, sheetName)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("读取失败: %v", err)), nil
	}

	// 转换为 JSON
	jsonData, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("JSON 序列化失败: %v", err)), nil
	}

	return mcp.NewToolResultText(string(jsonData)), nil
}

// handleWriteExcel 处理写入 Excel
func handleWriteExcel(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	filePath, err := request.RequireString("file_path")
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("file_path 参数错误: %v", err)), nil
	}

	dataStr, err := request.RequireString("data")
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("data 参数错误: %v", err)), nil
	}

	sheetName := request.GetString("sheet_name", "")

	log.Printf("write_excel 工具被调用: file_path=%s, sheet_name=%s", filePath, sheetName)

	// 解析 JSON 数据
	var data []map[string]interface{}
	if err := json.Unmarshal([]byte(dataStr), &data); err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("JSON 解析失败: %v", err)), nil
	}

	// 写入 Excel
	if err := excel.WriteExcel(filePath, data, sheetName); err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("写入失败: %v", err)), nil
	}

	return mcp.NewToolResultText(fmt.Sprintf("✅ 文件写入成功: %s", filePath)), nil
}
