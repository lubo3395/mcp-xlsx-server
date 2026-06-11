import { z } from "zod";

export const ListSheetsSchema = z.object({
  file_path: z.string().min(1, "文件路径不能为空").describe("xlsx 文件路径"),
}).strict();

export const ReadFormatEnum = z.enum(["json", "array"]);

export const ReadSheetSchema = z.object({
  file_path: z.string().min(1, "文件路径不能为空").describe("xlsx 文件路径"),
  sheet: z.union([
    z.string().min(1, "sheet 名称不能为空"),
    z.number().int().min(0, "sheet 索引从 0 开始"),
  ]).describe("sheet 名称或索引 (从 0 开始)"),
  format: ReadFormatEnum.default("json").describe(
    "输出格式: 'json' 返回对象数组 (第一行为表头), 'array' 返回二维数组"
  ),
  range: z.string().optional().describe(
    "可选单元格范围, 如 'A1:C10'"
  ),
  header_row: z.number().int().min(0).default(0).describe(
    "表头行号 (仅 json 格式有效), 从 0 开始, 默认第一行为表头"
  ),
}).strict();

export const CreateFileSchema = z.object({
  file_path: z.string().min(1, "文件路径不能为空").describe("输出 xlsx 文件路径"),
  data: z.string().min(1, "数据不能为空").describe(
    "JSON 数据: 可以是对象数组 [{\"col1\":\"val1\",...}] 或二维数组 [[\"h1\",\"h2\"],[\"v1\",\"v2\"]]"
  ),
  sheet_name: z.string().min(1).default("Sheet1").describe("sheet 名称"),
  overwrite: z.boolean().default(false).describe("如果文件已存在是否覆盖"),
}).strict();

export const AddSheetSchema = z.object({
  file_path: z.string().min(1, "文件路径不能为空").describe("已存在的 xlsx 文件路径"),
  data: z.string().min(1, "数据不能为空").describe(
    "JSON 数据: 可以是对象数组或二维数组"
  ),
  sheet_name: z.string().min(1, "sheet 名称不能为空").describe("新 sheet 名称 (不能与已有 sheet 重名)"),
}).strict();

export const UpdateCellsSchema = z.object({
  file_path: z.string().min(1, "文件路径不能为空").describe("xlsx 文件路径"),
  sheet: z.union([
    z.string().min(1),
    z.number().int().min(0),
  ]).describe("sheet 名称或索引 (从 0 开始)"),
  cells: z.string().min(1, "cells 不能为空").describe(
    "要更新的单元格, JSON 对象格式: {\"A1\": \"值\", \"B2\": 123}"
  ),
}).strict();
