# Excel MCP Server

一个基于 [Model Context Protocol (MCP)](https://github.com/mark3labs/mcp-go) 的 Excel/CSV 文件读写工具。支持 XLSX 和 CSV 格式的文件操作。

## 功能特性

- ✅ **读取 Excel/CSV 文件**：支持 `.xlsx`、`.xlsm`、`.xltx`、`.xltm` 和 `.csv` 格式
- ✅ **写入 Excel/CSV 文件**：支持多种 Excel 格式和 CSV 格式
- ✅ **多 Sheet 支持**：对于 Excel 文件，支持指定 Sheet 名称读写
- ✅ **JSON 格式数据**：输入输出均采用 `[]map[string]interface{}` 的 JSON 格式
- ✅ **自动类型识别**：自动识别数字和字符串类型

## 安装

```bash
git clone https://github.com/Xuzan9396/excel_mcp.git
cd excel_mcp
go mod tidy
go build -o excel-mcp cmd/excel-mcp/main.go
```

## 使用方法

### 启动 MCP Server

```bash
./excel-mcp
```

服务器将通过标准输入/输出（STDIO）与客户端通信。

### MCP 工具

#### 1. read_excel

读取 Excel 或 CSV 文件。

**参数：**
- `file_path` (必需)：文件路径，支持格式：`.xlsx`、`.xlsm`、`.xltx`、`.xltm`、`.csv`
- `sheet_name` (可选)：Sheet 名称，仅对 Excel 文件有效。为空则读取第一个 Sheet

**返回：**
JSON 格式的数据数组，格式为：
```json
[
  {"列名1": 值1, "列名2": 值2, "列名3": 值3},
  {"列名1": 值1, "列名2": 值2, "列名3": 值3}
]
```

**示例：**
```json
{
  "name": "read_excel",
  "arguments": {
    "file_path": "/path/to/file.xlsx",
    "sheet_name": "Sheet1"
  }
}
```

#### 2. write_excel

写入数据到 Excel 或 CSV 文件。

**参数：**
- `file_path` (必需)：文件路径，支持格式：`.xlsx`、`.xlsm`、`.xltx`、`.xltm`、`.csv`
- `data` (必需)：JSON 格式的数据字符串，格式为 `[{"列名1":值1,"列名2":值2}]`
- `sheet_name` (可选)：Sheet 名称，仅对 Excel 文件有效。为空则使用 "Sheet1"

**返回：**
成功消息字符串

**示例：**
```json
{
  "name": "write_excel",
  "arguments": {
    "file_path": "/path/to/output.xlsx",
    "data": "[{\"姓名\":\"张三\",\"年龄\":25},{\"姓名\":\"李四\",\"年龄\":30}]",
    "sheet_name": "员工信息"
  }
}
```

## 数据格式说明

### 输入格式（写入）

```json
[
  {
    "列名A": "值1",
    "列名B": 123,
    "列名C": "值2"
  },
  {
    "列名A": "值3",
    "列名B": 456,
    "列名C": "值4"
  }
]
```

写入时，字典的 key 将作为列标题，value 作为单元格值。

### 输出格式（读取）

读取时返回相同的格式，第一行作为列标题，后续行作为数据。

## 项目结构

```
excel_mcp/
├── cmd/
│   └── excel-mcp/
│       └── main.go              # MCP Server 主程序
├── internal/
│   └── excel/
│       └── excel.go             # Excel/CSV 读写核心功能
├── test_mcp/
│   ├── test_excel.go            # 功能测试
│   └── test_mcp_client.go       # MCP 客户端测试
├── go.mod
├── go.sum
├── .gitignore
└── README.md
```

## 测试

### 运行功能测试

```bash
go run test_mcp/test_excel.go
```

### 运行 MCP 客户端测试

```bash
go run test_mcp/test_mcp_client.go
```

## 依赖项

- [excelize/v2](https://github.com/xuri/excelize) - Excel 文件处理
- [mcp-go](https://github.com/mark3labs/mcp-go) - MCP 协议实现

## 技术细节

- **Excel 处理**：使用 `excelize/v2` 库处理 XLSX 格式文件
- **CSV 处理**：使用 Go 标准库 `encoding/csv` 处理 CSV 文件
- **MCP 协议**：基于 `mcp-go` 库实现 Model Context Protocol
- **传输方式**：STDIO（标准输入/输出）

## 许可证

MIT License

## 作者

Xuzan9396

## 贡献

欢迎提交 Issue 和 Pull Request！
