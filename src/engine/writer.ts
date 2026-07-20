// 对应 docs/03-detailed-design.md 第 2.18 节：src/engine/writer.ts 编辑适配

import * as fs from "node:fs";
import * as path from "node:path";
import ExcelJS from "exceljs";
import { ToolError } from "./errors.js";
import type { CellValue, BookFormat } from "./types.js";

/**
 * 对应第 2.18 节：
 * 新建一个 xlsx 工作簿并保存。
 * 文件已存在 → ToolError("FILE_EXISTS")
 */
export async function createWorkbook(
  filePath: string,
  sheetNames: string[]
): Promise<void> {
  if (fs.existsSync(filePath)) {
    throw new ToolError(
      "FILE_EXISTS",
      `文件已存在：${filePath}。如需修改请直接使用写入类工具；如需重建请先删除该文件。`
    );
  }

  const wb = new ExcelJS.Workbook();
  for (const name of sheetNames) {
    wb.addWorksheet(name);
  }

  await atomicSave(wb, filePath, "xlsx");
}

/**
 * 对应第 2.18 节：
 * 打开 → 执行 mutator → 原子保存。
 * - 文件不存在 → FILE_NOT_FOUND
 * - .xls → READ_ONLY_FORMAT（ExcelJS 不支持 xls）
 * - .xlsx/.xlsm → workbook.xlsx.readFile；.csv → workbook.csv.readFile
 * - mutator 内抛的 ToolError 原样向上抛；抛其他异常 → UNKNOWN
 */
export async function editWorkbook(
  filePath: string,
  mutator: (wb: ExcelJS.Workbook) => void | Promise<void>
): Promise<void> {
  if (!fs.existsSync(filePath)) {
    throw new ToolError(
      "FILE_NOT_FOUND",
      `文件不存在：${filePath}。请检查路径是否正确（需要绝对路径）。`
    );
  }

  const ext = path.extname(filePath).toLowerCase();
  let format: BookFormat;
  const wb = new ExcelJS.Workbook();

  try {
    if (ext === ".xls") {
      throw new ToolError(
        "READ_ONLY_FORMAT",
        "老版 .xls 文件不支持直接修改。请先在 Excel/WPS 中另存为 .xlsx 后再操作。"
      );
    }

    if (ext === ".xlsx" || ext === ".xlsm") {
      format = "xlsx";
      await wb.xlsx.readFile(filePath);
    } else if (ext === ".csv") {
      format = "csv";
      await wb.csv.readFile(filePath);
    } else {
      throw new ToolError(
        "UNSUPPORTED_FORMAT",
        `不支持的文件格式：${ext}。仅支持 .xlsx / .xls / .csv。`
      );
    }
  } catch (e) {
    if (e instanceof ToolError) throw e;
    throw new ToolError(
      "UNKNOWN",
      `无法解析文件：${filePath}，文件可能已损坏或不是有效的表格文件。`
    );
  }

  try {
    await mutator(wb);
  } catch (e) {
    if (e instanceof ToolError) throw e;
    throw new ToolError(
      "UNKNOWN",
      `操作失败：${e instanceof Error ? e.message : "未知错误"}`
    );
  }

  try {
    await atomicSave(wb, filePath, format);
  } catch (e) {
    if (e instanceof ToolError) throw e;
    throw new ToolError(
      "UNKNOWN",
      `操作失败：${e instanceof Error ? e.message : "未知错误"}`
    );
  }
}

/**
 * 对应第 2.18 节：
 * 内部函数，不导出。
 * 1. 写临时文件 <原名>.tmp-<pid>
 * 2. 成功后 fs.renameSync 替换原文件
 * 3. finally 清理临时文件
 */
async function atomicSave(
  wb: ExcelJS.Workbook,
  filePath: string,
  format: BookFormat
): Promise<void> {
  const tmpPath = `${filePath}.tmp-${process.pid}`;

  try {
    if (format === "csv") {
      await wb.csv.writeFile(tmpPath);
    } else {
      await wb.xlsx.writeFile(tmpPath);
    }
  } catch (e) {
    if (isNodeError(e) && e.code === "ENOENT") {
      throw new ToolError(
        "DIR_NOT_FOUND",
        `保存失败：目录不存在：${path.dirname(filePath)}。`
      );
    }
    throw e;
  }

  try {
    fs.renameSync(tmpPath, filePath);
  } catch (e) {
    if (isNodeError(e) && (e.code === "EPERM" || e.code === "EACCES" || e.code === "EBUSY")) {
      throw new ToolError(
        "FILE_BUSY",
        `无法保存：${filePath} 正被其他程序（如 Excel）占用，请关闭后重试。`
      );
    }
    throw e;
  } finally {
    try {
      if (fs.existsSync(tmpPath)) {
        fs.unlinkSync(tmpPath);
      }
    } catch {
      // 尽力清理，忽略错误
    }
  }
}

function isNodeError(e: unknown): e is NodeJS.ErrnoException {
  return e instanceof Error && "code" in e;
}

/**
 * 对应第 2.18 节：
 * ExcelJS 原始 cell.value → 统一 CellValue（sort/dedupe 比较用）。
 */
export function normalizeCellValue(v: ExcelJS.CellValue): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === "string") return v;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString();

  if (typeof v === "object") {
    // 富文本：拼接各段 text
    if ("richText" in v && Array.isArray(v.richText)) {
      return v.richText.map((part) => String(part.text ?? "")).join("");
    }

    // 超链接：取 text
    if ("hyperlink" in v && "text" in v) {
      return String(v.text);
    }

    // 公式或共享公式：取 result 并继续归一化
    if ("formula" in v || "sharedFormula" in v) {
      const result = "result" in v ? (v as any).result : undefined;
      return normalizeCellValue(result as ExcelJS.CellValue);
    }

    // 错误值
    if ("error" in v) {
      return null;
    }
  }

  return null;
}
