import XLSX from "xlsx";
import path from "path";
import type { CreateResult, AddSheetResult, UpdateCellsResult } from "../types.js";
import {
  openWorkbook,
  getSheet,
  getSheetInfo,
  listSheets,
  parseDataInput,
  handleError,
  checkOutputPath,
} from "../utils.js";

/** 将解析后的数据转换为 worksheet */
function dataToSheet(data: Record<string, unknown>[] | unknown[][]): XLSX.WorkSheet {
  if (Array.isArray(data[0])) {
    return XLSX.utils.aoa_to_sheet(data as unknown[][]);
  }
  return XLSX.utils.json_to_sheet(data as Record<string, unknown>[]);
}

/** 统计数据的行列数 */
function getDataDimensions(data: Record<string, unknown>[] | unknown[][]): { rows: number; cols: number } {
  if (data.length === 0) return { rows: 0, cols: 0 };
  if (Array.isArray(data[0])) {
    const arr = data as unknown[][];
    return { rows: arr.length, cols: Math.max(...arr.map((r) => r.length)) };
  }
  const objs = data as Record<string, unknown>[];
  return { rows: objs.length, cols: Object.keys(objs[0]).length };
}

// ===== Tool Handlers =====

export async function handleCreateFile(
  filePath: string,
  data: string,
  sheetName: string,
  overwrite: boolean
): Promise<CreateResult> {
  const resolvedPath = checkOutputPath(filePath, overwrite);

  const parsed = parseDataInput(data);
  const ws = dataToSheet(parsed);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, resolvedPath);

  const { rows, cols } = getDataDimensions(parsed);
  return { file_path: resolvedPath, sheet_name: sheetName, rows, cols };
}

export async function handleAddSheet(
  filePath: string,
  data: string,
  sheetName: string
): Promise<AddSheetResult> {
  const resolvedPath = path.resolve(filePath);
  const wb = openWorkbook(resolvedPath);

  // 检查 sheet 名称是否已存在
  if (wb.SheetNames.includes(sheetName)) {
    throw new Error(
      `Sheet "${sheetName}" 已存在。可用的 sheet: ${wb.SheetNames.join(", ")}\n请使用不同的 sheet 名称。`
    );
  }

  const parsed = parseDataInput(data);
  const ws = dataToSheet(parsed);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, resolvedPath);

  const { rows, cols } = getDataDimensions(parsed);
  const sheets = listSheets(wb);
  return {
    file_path: resolvedPath,
    sheet_name: sheetName,
    rows,
    cols,
    sheets: sheets.map((s) => s.name),
  };
}

export async function handleUpdateCells(
  filePath: string,
  sheet: string | number,
  cellsStr: string
): Promise<UpdateCellsResult> {
  const resolvedPath = path.resolve(filePath);
  const wb = openWorkbook(resolvedPath);
  const ws = getSheet(wb, sheet);

  let sheetName: string;
  if (typeof sheet === "string") {
    sheetName = sheet;
  } else {
    sheetName = wb.SheetNames[sheet];
  }

  // 解析 cells JSON
  let cells: Record<string, unknown>;
  try {
    cells = JSON.parse(cellsStr);
    if (typeof cells !== "object" || cells === null || Array.isArray(cells)) {
      throw new Error("cells 必须是对象格式");
    }
  } catch (e) {
    if (e instanceof SyntaxError) {
      throw new Error(
        `cells 格式无效: 不是合法的 JSON 字符串\n正确格式: {"A1": "值", "B2": 123}`
      );
    }
    throw e;
  }

  if (Object.keys(cells).length === 0) {
    throw new Error("cells 不能为空对象。至少需要指定一个单元格，如: {\"A1\": \"值\"}");
  }

  // 验证并更新单元格
  const cellRefRegex = /^[A-Za-z]+[1-9]\d*$/;
  for (const cellRef of Object.keys(cells)) {
    if (!cellRefRegex.test(cellRef)) {
      throw new Error(
        `单元格引用无效: "${cellRef}"。正确格式如 "A1", "B2", "Z10"`
      );
    }

    const value = cells[cellRef];
    const cell: XLSX.CellObject = { t: "s", v: String(value) };

    if (typeof value === "number") {
      cell.t = "n";
      cell.v = value;
    } else if (typeof value === "boolean") {
      cell.t = "b";
      cell.v = value;
    } else if (value === null) {
      cell.t = "z";
      cell.v = undefined;
    } else {
      cell.t = "s";
      cell.v = String(value);
    }

    ws[cellRef] = cell;
  }

  XLSX.writeFile(wb, resolvedPath);

  return {
    file_path: resolvedPath,
    sheet: sheetName,
    updated: Object.keys(cells).length,
  };
}

// ===== Format Functions =====

export function formatCreateFile(result: CreateResult): string {
  return [
    `## 文件创建成功`,
    `- 路径: ${result.file_path}`,
    `- Sheet: ${result.sheet_name}`,
    `- 大小: ${result.rows} 行 × ${result.cols} 列`,
  ].join("\n");
}

export function formatAddSheet(result: AddSheetResult): string {
  return [
    `## Sheet 添加成功`,
    `- 文件: ${result.file_path}`,
    `- 新 Sheet: ${result.sheet_name}`,
    `- 大小: ${result.rows} 行 × ${result.cols} 列`,
    `- 所有 Sheet: ${result.sheets.join(", ")}`,
  ].join("\n");
}

export function formatUpdateCells(result: UpdateCellsResult): string {
  return [
    `## 单元格更新成功`,
    `- 文件: ${result.file_path}`,
    `- Sheet: ${result.sheet}`,
    `- 更新单元格数: ${result.updated}`,
  ].join("\n");
}

// ===== Tool Registration Configs =====

import { CreateFileSchema, AddSheetSchema, UpdateCellsSchema } from "../schemas.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export function registerWriteTools(server: McpServer): void {
  server.registerTool(
    "xlsx_create",
    {
      title: "创建 Excel 文件",
      description: `从 JSON 数据创建新的 xlsx 文件。

参数:
  - file_path (string, 必填): 输出 xlsx 文件路径
  - data (string, 必填): JSON 数据字符串
    - 对象数组格式: [{"姓名": "张三", "年龄": 30}, {"姓名": "李四", "年龄": 25}]
    - 二维数组格式: [["姓名", "年龄"], ["张三", 30], ["李四", 25]]
  - sheet_name (string, 默认 "Sheet1"): sheet 名称
  - overwrite (boolean, 默认 false): 文件已存在时是否覆盖

返回:
  - file_path: 输出文件路径
  - sheet_name: sheet 名称
  - rows: 总行数
  - cols: 总列数

示例:
  - "创建 test.xlsx，包含 [{name:'张三',age:30},{name:'李四',age:25}]" -> { file_path: "test.xlsx", data: '[{"name":"张三","age":30},{"name":"李四","age":25}]' }
  - "覆盖已有文件" -> { file_path: "test.xlsx", data: "...", overwrite: true }

错误处理:
  - 文件已存在且未设置 overwrite: "文件已存在: ..."
  - 输出目录不存在: "输出目录不存在: ..."
  - 数据格式无效: "数据格式无效: ..."`,
      inputSchema: CreateFileSchema,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await handleCreateFile(
          params.file_path,
          params.data,
          params.sheet_name,
          params.overwrite
        );
        return {
          content: [{ type: "text", text: formatCreateFile(result) }],
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
    "xlsx_add_sheet",
    {
      title: "添加 Sheet 到 Excel",
      description: `向已有的 xlsx 文件中添加新的 sheet。

参数:
  - file_path (string, 必填): 已存在的 xlsx 文件路径
  - data (string, 必填): JSON 数据字符串 (同 xlsx_create 的 data 参数)
  - sheet_name (string, 必填): 新 sheet 名称 (不能与已有 sheet 重名)

返回:
  - file_path: 文件路径
  - sheet_name: 新 sheet 名称
  - rows: 数据行数
  - cols: 数据列数
  - sheets: 更新后的所有 sheet 名称列表

示例:
  - "往 test.xlsx 添加一个名为 成绩表 的 sheet" -> { file_path: "test.xlsx", data: '[{"姓名":"张三","成绩":90}]', sheet_name: "成绩表" }

错误处理:
  - Sheet 名称已存在: "Sheet "名称" 已存在。可用 ..."
  - 文件不存在: "文件不存在: ..."`,
      inputSchema: AddSheetSchema,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await handleAddSheet(
          params.file_path,
          params.data,
          params.sheet_name
        );
        return {
          content: [{ type: "text", text: formatAddSheet(result) }],
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
    "xlsx_update_cells",
    {
      title: "更新 Excel 单元格",
      description: `更新 xlsx 文件中指定 sheet 的单元格值。

参数:
  - file_path (string, 必填): xlsx 文件路径
  - sheet (string|number, 必填): sheet 名称或索引 (从 0 开始)
  - cells (string, 必填): 要更新的单元格，JSON 对象格式: {"A1": "值", "B2": 123, "C3": true}

返回:
  - file_path: 文件路径
  - sheet: 更新的 sheet
  - updated: 更新单元格数量

示例:
  - "把 test.xlsx 的 Sheet1 中 A1 改成 Hello" -> { file_path: "test.xlsx", sheet: "Sheet1", cells: '{"A1": "Hello"}' }
  - "批量修改多个单元格" -> { file_path: "test.xlsx", sheet: 0, cells: '{"A1": "标题", "B1": "值"}' }

错误处理:
  - cells 格式无效: "cells 格式无效: ..."
  - 单元格引用无效: "单元格引用无效: ..."
  - sheet 不存在: "Sheet ... 不存在"`,
      inputSchema: UpdateCellsSchema,
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const result = await handleUpdateCells(
          params.file_path,
          params.sheet,
          params.cells
        );
        return {
          content: [{ type: "text", text: formatUpdateCells(result) }],
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
