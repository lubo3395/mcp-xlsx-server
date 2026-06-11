# mcp-xlsx-server

基于 [SheetJS (xlsx)](https://www.npmjs.com/package/xlsx) 的 Excel 文件读写 MCP 服务。

## 功能

提供 5 个 MCP 工具：

| 工具 | 功能 |
|------|------|
| `xlsx_list_sheets` | 列出 xlsx 文件中所有 sheet 名称和行列数 |
| `xlsx_read` | 读取指定 sheet 数据，支持 JSON/二维数组格式和范围选择 |
| `xlsx_create` | 从 JSON 数据创建新的 xlsx 文件 |
| `xlsx_add_sheet` | 向已有 xlsx 文件添加新 sheet |
| `xlsx_update_cells` | 更新指定单元格值 |

## 安装

```bash
npm install -g mcp-xlsx-server
```

或直接运行：

```bash
npx mcp-xlsx-server
```

## 配置

### 在 Claude Code 中

编辑 `claude.json`，使用 `mcpServers` 字段：

```json
{
  "mcpServers": {
    "xlsx": {
      "command": "npx",
      "args": ["mcp-xlsx-server"]
    }
  }
}
```

### 在 VS Code (GitHub Copilot) 中

编辑 `.vscode/mcp.json` 或全局 `mcp.json`，使用 **`servers`** 字段（不是 `mcpServers`）：

```json
{
  "servers": {
    "xlsx": {
      "type": "stdio",
      "command": "npx",
      "args": ["mcp-xlsx-server"]
    }
  }
}
```

文件位置：
- **工作区级别**: `.vscode/mcp.json`
- **全局**: `%APPDATA%\Code\User\mcp.json`

### 在 Cline / Roo Code (VS Code 扩展) 中

编辑 `cline.json` 或 `mcp.json`，使用 `mcpServers` 字段：

```json
{
  "mcpServers": {
    "xlsx": {
      "command": "npx",
      "args": ["-y", "mcp-xlsx-server"]
    }
  }
}
```

### 在 Cursor 中

编辑 `.cursor/mcp.json`，使用 `mcpServers` 字段：

```json
{
  "mcpServers": {
    "xlsx": {
      "command": "npx",
      "args": ["mcp-xlsx-server"]
    }
  }
}
```

## 工具详解

### xlsx_list_sheets

列出 xlsx 文件中所有 sheet。

参数：
- `file_path` (string, 必填): xlsx 文件路径

### xlsx_read

读取指定 sheet 的数据。

参数：
- `file_path` (string, 必填): xlsx 文件路径
- `sheet` (string|number, 必填): sheet 名称或索引 (从 0 开始)
- `format` (string, 默认 "json"): "json" 返回对象数组, "array" 返回二维数组
- `range` (string, 可选): 单元格范围，如 `"A1:C10"`

### xlsx_create

从 JSON 数据创建新的 xlsx 文件。

参数：
- `file_path` (string, 必填): 输出路径
- `data` (string, 必填): JSON 数据字符串
- `sheet_name` (string, 默认 "Sheet1"): sheet 名称
- `overwrite` (boolean, 默认 false): 是否覆盖已有文件

### xlsx_add_sheet

向已有 xlsx 文件添加新 sheet。

参数：
- `file_path` (string, 必填): 已存在的 xlsx 文件路径
- `data` (string, 必填): JSON 数据字符串
- `sheet_name` (string, 必填): 新 sheet 名称

### xlsx_update_cells

更新指定单元格。

参数：
- `file_path` (string, 必填): xlsx 文件路径
- `sheet` (string|number, 必填): sheet 名称或索引
- `cells` (string, 必填): JSON 对象，如 `{"A1": "值", "B2": 123}`

## 本地开发

```bash
# 克隆仓库
git clone https://github.com/lubo3395/mcp-xlsx-server.git
cd mcp-xlsx-server

# 安装依赖
npm install

# 开发模式 (热重载)
npm run dev

# 构建
npm run build

# 运行
npm start

# MCP Inspector 测试
npx @modelcontextprotocol/inspector node dist/index.js
```

## License

MIT
