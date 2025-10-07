# @xuzan/excel-mcp

Excel MCP Server - Excel/CSV 文件读写 MCP 服务器

## 快速开始

### 使用 npx（推荐）

```bash
npx -y @xuzan/excel-mcp
```

### 配置到 Claude Desktop

在 Claude Desktop 配置文件中添加：

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": ["-y", "@xuzan/excel-mcp"]
    }
  }
}
```

## 功能特性

- ✅ **读取 Excel**: 支持 .xlsx, .xlsm, .xltx, .xltm 格式
- ✅ **读取 CSV**: 支持 CSV 文件读取
- ✅ **写入 Excel**: 支持将 JSON 数据写入 Excel/CSV
- ✅ **跨平台**: macOS / Linux / Windows 全平台支持
- ✅ **简单易用**: MCP 协议，与 Claude 无缝集成

## 支持的平台

- macOS (Apple Silicon / Intel)
- Linux (x64)
- Windows (x64 / ARM64)

## 使用说明

在 Claude Desktop 中直接对话：

```
帮我读取 /path/to/file.xlsx 文件
```

或者写入数据：

```
帮我将这些数据写入到 output.xlsx 文件
```

## 完整文档

查看 [GitHub 仓库](https://github.com/Xuzan9396/excel_mcp) 了解更多信息。
git add internal/excel/excel.go && git commit -m "fix: 修复写入时大数字转科学计数法问题

      - CSV 写入：添加 formatValue 函数，正确格式化不同类型的数值
      - Excel 写入：为数字单元格设置正确的数字格式，防止科学计数法
      - 支持所有整数类型（int, int8-64, uint, uint8-64）
      - 支持浮点数类型（float32, float64），整数值显示为整数
      - 测试通过：3500364 和 900490861 正确显示，不再转换为 3.50E+06 和 9.00E+08"

git push

git tag -a v0.0.2 -m "Release v0.0.2 - 修复大数字转科学计数法问题" && git push origin v0.0.2

## 许可证

MIT
