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

## 问题反馈

如有问题，请在 [GitHub Issues](https://github.com/Xuzan9396/excel_mcp/issues) 中反馈。

## 许可证

MIT
