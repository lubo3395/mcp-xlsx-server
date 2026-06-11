import XLSX from "xlsx";
import type { ListSheetsResult, ReadResult } from "../types.js";
import { openWorkbook, getSheet, getSheetInfo, listSheets, handleError } from "../utils.js";

/** 获取 worksheet 中指定范围 */
function getRangeString(range?: string, ws?: XLSX.WorkSheet): XLSX.Range | undefined {
  if (range) {
    try {
      return XLSX.utils.decode_range(range);
    } catch {
      throw new Error(
        `范围格式无效: "${range}"。正确格式如 "A1:C10"`
      );
    }
  }
  return undefined;
}

/** 读取 sheet 数据为对象数组 (第一行为表头) */
function readAsJson(ws: XLSX.WorkSheet, range?: string): Record<string, unknown>[] {
  const opts: XLSX.Sheet2JSONOpts = { defval: null };
  const rangeObj = getRangeString(range);
  if (rangeObj) {
    opts.range = rangeObj;
  }
  return XLSX.utils.sheet_to_json(ws, opts);
}

/** 读取 sheet 数据为二维数组 */
function readAsArray(ws: XLSX.WorkSheet, range?: string): unknown[][] {
  const opts: XLSX.Sheet2JSONOpts = { header: 1, defval: null };
  const rangeObj = getRangeString(range);
  if (rangeObj) {
    opts.range = rangeObj;
  }
  return XLSX.utils.sheet_to_json(ws, opts);
}

/** 将表格数据格式化为 Markdown 表格 */
function formatAsMarkdownTable(data: Record<string, unknown>[] | unknown[][]): string {
  if (data.length === 0) return "(空数据)";

  const isJson = !Array.isArray(data[0]);
  let headers: string[];
  let rows: unknown[][];

  if (isJson) {
    const jsonData = data as Record<string, unknown>[];
    headers = Object.keys(jsonData[0]);
    rows = jsonData.map((row) => headers.map((h) => row[h]));
  } else {
    const arrayData = data as unknown[][];
    if (arrayData.length === 0) return "(空数据)";
    headers = (arrayData[0] as string[]).map((h) => String(h ?? ""));
    rows = arrayData.slice(1);
  }

  const cols = headers.length;
  if (cols === 0) return "(无列数据)";

  // 计算每列最大宽度
  const colWidths = headers.map((h, i) => {
    const headerLen = String(h).length;
    const maxDataLen = rows.reduce((max, row) => {
      const val = i < row.length ? String(row[i] ?? "") : "";
      return Math.max(max, val.length);
    }, 0);
    return Math.max(headerLen, maxDataLen, 3);
  });

  // 构建表格
  const lines: string[] = [];

  // 表头
  const headerRow = "| " + headers.map((h, i) => String(h).padEnd(colWidths[i])).join(" | ") + " |";
  lines.push(headerRow);

  // 分隔线
  const sepRow = "| " + colWidths.map((w) => "-".repeat(w)).join(" | ") + " |";
  lines.push(sepRow);

  // 数据行 (最多 50 行)
  const displayRows = rows.slice(0, 50);
  for (const row of displayRows) {
    const rowStr = "| " + colWidths.map((w, i) => {
      const val = i < row.length ? String(row[i] ?? "") : "";
      return val.padEnd(w);
    }).join(" | ") + " |";
    lines.push(rowStr);
  }

  if (rows.length > 50) {
    lines.push(`\n*仅显示前 50 行，共 ${rows.length} 行*`);
  }

  return lines.join("\n");
}

// ===== Tool Handlers =====

export async function handleListSheets(filePath: string): Promise<ListSheetsResult> {
  const wb = openWorkbook(filePath);
  const sheets = listSheets(wb);
  return { file_path: filePath, sheets };
}

export async function handleReadSheet(
  filePath: string,
  sheet: string | number,
  format: "json" | "array",
  range?: string
): Promise<ReadResult> {
  const wb = openWorkbook(filePath);
  const ws = getSheet(wb, sheet);

  let sheetName: string;
  if (typeof sheet === "string") {
    sheetName = sheet;
  } else {
    sheetName = wb.SheetNames[sheet];
  }

  const { rows, cols } = getSheetInfo(ws);

  let data: Record<string, unknown>[] | unknown[][];
  if (format === "json") {
    data = readAsJson(ws, range);
  } else {
    data = readAsArray(ws, range);
  }

  return { sheet: sheetName, rows: data.length, cols: format === "json" ? Object.keys(data[0] || {}).length : cols, data, format };
}

// ===== Format Functions =====

export function formatListSheets(result: ListSheetsResult): string {
  if (result.sheets.length === 0) {
    return `文件 "${result.file_path}" 中没有 sheet。`;
  }

  const lines: string[] = [
    `## 文件: ${result.file_path}\n`,
    `共 ${result.sheets.length} 个 sheet:\n`,
  ];

  for (const s of result.sheets) {
    lines.push(`- **${s.name}**: ${s.rows} 行 × ${s.cols} 列`);
  }

  return lines.join("\n");
}

export function formatReadSheet(result: ReadResult): string {
  const lines: string[] = [
    `## Sheet: ${result.sheet}`,
    `行数: ${result.rows}, 列数: ${result.cols}\n`,
  ];

  if (result.rows === 0) {
    lines.push("(空 sheet)");
  } else {
    const markdown = formatAsMarkdownTable(result.data);
    lines.push(markdown);
  }

  return lines.join("\n");
}

// ===== Tool Registration Configs =====

import { ListSheetsSchema, ReadSheetSchema } from "../schemas.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export function registerReadTools(server: McpServer): void {
  server.registerTool(
    "xlsx_list_sheets",
    {
      title: "列出 Excel Sheet",
      description: `列出 xlsx 文件中所有 sheet 的名称、行数和列数。

参数:
  - file_path (string, 必填): xlsx 文件路径

返回:
  - file_path: 文件路径
  - sheets: Sheet 信息数组
    - name: Sheet 名称
    - rows: 行数
    - cols: 列数

示例:
  - "读取 test.xlsx 有哪些 sheet" -> { file_path: "test.xlsx" }

错误处理:
  - 文件不存在: "文件不存在: <路径>"
  - 不是 xlsx 文件: "不支持的文件格式: <路径>"`,
      inputSchema: ListSheetsSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await handleListSheets(params.file_path);
        return {
          content: [{ type: "text", text: formatListSheets(result) }],
          structuredContent: result as unknown as Record<string, unknown>,
        };
      } catch (error) {
        return {
          isError: true,
          content: [{ type: "text", text: handleError(error) }],
        };
      }
    }
  );

  server.registerTool(
    "xlsx_read",
    {
      title: "读取 Excel Sheet 数据",
      description: `读取 xlsx 文件中指定 sheet 的数据，支持 JSON 对象数组和二维数组两种格式。

参数:
  - file_path (string, 必填): xlsx 文件路径
  - sheet (string|number, 必填): sheet 名称或索引 (从 0 开始)
  - format (string, 默认 "json"): 输出格式
    - "json": 返回对象数组 (第一行为表头)
    - "array": 返回二维数组
  - range (string, 可选): 单元格范围，如 "A1:C10"
  - header_row (number, 默认 0): 表头行号，仅 json 格式有效

返回:
  - sheet: 读取的 sheet 名称
  - rows: 数据行数
  - cols: 数据列数
  - data: 表格数据
  - format: 数据格式

示例:
  - "读取 test.xlsx 的 Sheet1" -> { file_path: "test.xlsx", sheet: "Sheet1" }
  - "读取 test.xlsx 第一个 sheet 的 A1 到 C10 范围" -> { file_path: "test.xlsx", sheet: 0, range: "A1:C10" }
  - "以二维数组格式读取" -> { file_path: "test.xlsx", sheet: "Sheet1", format: "array" }

错误处理:
  - sheet 不存在: "Sheet "名称" 不存在。可用的 sheet: ..."
  - sheet 索引无效: "Sheet 索引无效: ..."
  - 文件不存在: "文件不存在: <路径>"`,
      inputSchema: ReadSheetSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await handleReadSheet(
          params.file_path,
          params.sheet,
          params.format,
          params.range
        );
        return {
          content: [{ type: "text", text: formatReadSheet(result) }],
          structuredContent: result as unknown as Record<string, unknown>,
        };
      } catch (error) {
        return {
          isError: true,
          content: [{ type: "text", text: handleError(error) }],
        };
      }
    }
  );
}
