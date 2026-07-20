// 对应 docs/03-detailed-design.md 第 2.17 节：src/engine/reader.ts 读取适配

import * as fs from "node:fs";
import * as path from "node:path";
import XLSX from "xlsx";
import { ToolError } from "./errors.js";
import type { CellValue, UnifiedSheet, UnifiedWorkbook, BookFormat } from "./types.js";

/**
 * 对应第 2.17 节：
 * 用 SheetJS 把任意支持的文件读成 UnifiedWorkbook。
 * 全项目唯一 import xlsx 的地方。
 */
export function readWorkbook(filePath: string): UnifiedWorkbook {
  // 对应第 2.17 节：文件不存在 → FILE_NOT_FOUND（文案含路径，提示需要绝对路径）
  if (!fs.existsSync(filePath)) {
    throw new ToolError(
      "FILE_NOT_FOUND",
      `文件不存在：${filePath}。请检查路径是否正确（需要绝对路径）。`
    );
  }

  // 对应第 2.17 节：按扩展名判定格式
  const ext = path.extname(filePath).toLowerCase();
  let format: BookFormat;
  if (ext === ".xlsx" || ext === ".xlsm") {
    format = "xlsx";
  } else if (ext === ".xls") {
    format = "xls";
  } else if (ext === ".csv") {
    format = "csv";
  } else {
    throw new ToolError(
      "UNSUPPORTED_FORMAT",
      `不支持的文件格式：${ext}。仅支持 .xlsx / .xls / .csv。`
    );
  }

  // 对应第 2.17 节：XLSX.readFile 解析文件，抛异常 → 包装为 UNKNOWN
  let workbook;
  try {
    workbook = XLSX.readFile(filePath, { cellDates: true });
  } catch {
    throw new ToolError(
      "UNKNOWN",
      `无法解析文件：${filePath}，文件可能已损坏或不是有效的表格文件。`
    );
  }

  // 对应第 2.17 节：逐个 sheet 转统一内存结构
  const sheets: UnifiedSheet[] = workbook.SheetNames.map((name) => {
    const ws = workbook.Sheets[name];
    const rawRows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: true,
      defval: null,
    }) as unknown[][];

    const normalizedRows: CellValue[][] = rawRows.map((row) =>
      row.map((cell) => normalizeValue(cell))
    );

    const colCount = normalizedRows.reduce(
      (max, row) => Math.max(max, row.length),
      0
    );

    // 对应第 2.17 节：每行补齐到 colCount
    const paddedRows = normalizedRows.map((row) => {
      const padded = [...row];
      while (padded.length < colCount) {
        padded.push(null);
      }
      return padded;
    });

    // 对应第 2.17 节：裁掉尾部全空行（整行皆 null 或空数组）
    const trimmedRows = trimTrailingEmptyRows(paddedRows);
    const rowCount = trimmedRows.length;

    return {
      name,
      rowCount,
      colCount: rowCount > 0 ? colCount : 0,
      values: trimmedRows,
    };
  });

  return { format, sheets };
}

/**
 * 对应第 2.17 节：值的归一化规则
 * - Date → ISO 字符串
 * - string / number / boolean 原样保留
 * - 其他一切（undefined、富文本、错误值等）→ null
 * 公式格的缓存计算结果若 SheetJS 已返回为 number/string/boolean，则自然保留；
 * 无缓存结果时通常为 null 或对象，统一归 null。
 */
function normalizeValue(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === "string") return v;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v;
  if (v instanceof Date) return v.toISOString();
  return null;
}

/**
 * 对应第 2.17 节：裁掉尾部全空行
 * 整行皆 null 或空数组视为空行。
 */
function trimTrailingEmptyRows(rows: CellValue[][]): CellValue[][] {
  let last = rows.length - 1;
  while (last >= 0) {
    const row = rows[last];
    if (row.length === 0 || row.every((v) => v === null)) {
      last--;
    } else {
      break;
    }
  }
  return rows.slice(0, last + 1);
}
