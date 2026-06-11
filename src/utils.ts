import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import type { SheetInfo } from "./types.js";

/** 安全获取 sheet 对象 */
export function getSheet(wb: XLSX.WorkBook, sheet: string | number): XLSX.WorkSheet {
  let sheetName: string;
  if (typeof sheet === "number") {
    if (sheet < 0 || sheet >= wb.SheetNames.length) {
      throw new Error(
        `Sheet 索引无效: ${sheet}, 文件中共有 ${wb.SheetNames.length} 个 sheet (索引 0-${wb.SheetNames.length - 1})`
      );
    }
    sheetName = wb.SheetNames[sheet];
  } else {
    if (!wb.Sheets[sheet]) {
      throw new Error(
        `Sheet "${sheet}" 不存在。可用的 sheet: ${wb.SheetNames.join(", ")}`
      );
    }
    sheetName = sheet;
  }
  return wb.Sheets[sheetName];
}

/** 获取 sheet 基本信息 (行列数) */
export function getSheetInfo(ws: XLSX.WorkSheet): { rows: number; cols: number } {
  if (!ws["!ref"]) return { rows: 0, cols: 0 };
  const range = XLSX.utils.decode_range(ws["!ref"]);
  return {
    rows: range.e.r - range.s.r + 1,
    cols: range.e.c - range.s.c + 1,
  };
}

/** 获取 workbook 中所有 sheet 的信息 */
export function listSheets(wb: XLSX.WorkBook): SheetInfo[] {
  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const { rows, cols } = getSheetInfo(ws);
    return { name, rows, cols };
  });
}

/** 打开 xlsx 文件，返回 workbook */
export function openWorkbook(filePath: string): XLSX.WorkBook {
  const resolvedPath = path.resolve(filePath);

  if (!fs.existsSync(resolvedPath)) {
    throw new Error(`文件不存在: ${resolvedPath}`);
  }

  const stat = fs.statSync(resolvedPath);
  if (!stat.isFile()) {
    throw new Error(`路径不是文件: ${resolvedPath}`);
  }

  if (!resolvedPath.match(/\.xlsx?$/i)) {
    throw new Error(`不支持的文件格式: ${resolvedPath}, 仅支持 .xls 和 .xlsx`);
  }

  return XLSX.readFile(resolvedPath, { cellDates: true });
}

/** 解析数据参数: 支持 JSON 对象数组 或 二维数组 */
export function parseDataInput(data: string): Record<string, unknown>[] | unknown[][] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(data);
  } catch {
    throw new Error("数据格式无效: 不是合法的 JSON 字符串");
  }

  if (!Array.isArray(parsed)) {
    throw new Error("数据格式无效: 需要 JSON 数组 (对象数组或二维数组)");
  }

  if (parsed.length === 0) {
    throw new Error("数据不能为空数组");
  }

  return parsed;
}

/** 错误处理 */
export function handleError(error: unknown): string {
  if (error instanceof Error) {
    return `错误: ${error.message}`;
  }
  return `错误: ${String(error)}`;
}

/** 检查输出路径是否可写 */
export function checkOutputPath(filePath: string, overwrite: boolean): string {
  const resolvedPath = path.resolve(filePath);

  if (fs.existsSync(resolvedPath)) {
    if (!overwrite) {
      throw new Error(
        `文件已存在: ${resolvedPath}\n请设置 overwrite=true 覆盖，或使用不同的文件路径`
      );
    }
    const stat = fs.statSync(resolvedPath);
    if (stat.isDirectory()) {
      throw new Error(`路径是目录，不是文件: ${resolvedPath}`);
    }
  }

  const dir = path.dirname(resolvedPath);
  if (!fs.existsSync(dir)) {
    throw new Error(`输出目录不存在: ${dir}`);
  }

  return resolvedPath;
}
