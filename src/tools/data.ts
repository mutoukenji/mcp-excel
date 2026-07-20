// 对应 docs/03-detailed-design.md 第 2.12 节：src/tools/data.ts 数据工具

import * as path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { filePathSchema, rangeSchema, tool } from "./common.js";
import { readWorkbook } from "../engine/reader.js";
import { editWorkbook, normalizeCellValue } from "../engine/writer.js";
import { parseRange, resolveColumn } from "../engine/address.js";
import { ToolError } from "../engine/errors.js";
import type { CellValue } from "../engine/types.js";

// 对应第 2.12 节：filter_range 的条件 schema
const conditionSchema = z.object({
  column: z.string().min(1), // 列字母或表头名
  op: z.enum(["=", "!=", ">", ">=", "<", "<=", "contains"]),
  value: z.union([z.string(), z.number(), z.boolean()]),
});

/**
 * 对应第 2.12 节：
 * 导出 registerDataTools(server: McpServer): void
 * 注册 3 个数据工具：sort_range / filter_range / dedupe_range
 */
export function registerDataTools(server: McpServer): void {
  // 对应第 2.12 节 (1)：sort_range —— 排序
  server.registerTool(
    "sort_range",
    {
      title: "排序区域",
      description:
        "对一个区域按某列排序（会修改文件）。区域外的内容不动。有表头时传 hasHeader: true 并用表头名指定列更稳妥。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        range: rangeSchema,
        keyColumn: z.string().min(1),
        order: z.enum(["asc", "desc"]),
        hasHeader: z.boolean().optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const rangeInput = args.range as string;
      const keyColumn = args.keyColumn as string;
      const order = args.order as "asc" | "desc";
      const hasHeader = (args.hasHeader as boolean | undefined) ?? false;

      const addr = parseRange(rangeInput);

      await editWorkbook(filePath, (wb) => {
        // CSV 仅单工作表，使用第一张工作表
        const ws =
          path.extname(filePath).toLowerCase() === ".csv"
            ? wb.worksheets[0]
            : wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }

        // 对应第 2.12 节：读出原始 cell.value（保住日期/数字类型）与归一化值
        const rawRows: any[] = [];
        const normRows: CellValue[][] = [];
        for (let r = addr.start.row; r <= addr.end.row; r++) {
          const rawRow: any[] = [];
          const normRow: CellValue[] = [];
          for (let c = addr.start.col; c <= addr.end.col; c++) {
            const cell = ws.getCell(r, c);
            rawRow.push(cell.value);
            normRow.push(normalizeCellValue(cell.value));
          }
          rawRows.push(rawRow);
          normRows.push(normRow);
        }

        const headerRow = hasHeader ? normRows[0] : undefined;
        const dataStart = hasHeader ? 1 : 0;
        const dataRaw = rawRows.slice(dataStart);
        const dataNorm = normRows.slice(dataStart);

        // 对应第 2.12 节：解析 keyColumn → 区域内列下标
        const absCol = resolveColumn(keyColumn, addr, headerRow);
        const keyColIdx = absCol - addr.start.col;

        // 对应第 2.12 节：按 compareValues 稳定排序（V8 Array.prototype.sort 是稳定排序）
        const sortedRaw = dataRaw
          .map((row, idx) => ({ row, norm: dataNorm[idx] }))
          .sort((a, b) =>
            compareValues(a.norm[keyColIdx], b.norm[keyColIdx], order)
          )
          .map((item) => item.row);

        // 对应第 2.12 节：把 raw 按新顺序写回原区域
        for (let i = 0; i < rawRows.length; i++) {
          const srcRow = hasHeader && i === 0 ? rawRows[0] : sortedRaw[i - dataStart];
          for (let j = 0; j < srcRow.length; j++) {
            ws.getCell(addr.start.row + i, addr.start.col + j).value = srcRow[j] as any;
          }
        }
      });

      return {};
    })
  );

  // 对应第 2.12 节 (2)：filter_range —— 按条件筛选（只读，不改文件）
  server.registerTool(
    "filter_range",
    {
      title: "筛选区域",
      description:
        "按条件筛选区域中的行并返回结果，不修改文件。多个条件为'并且'关系。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        range: rangeSchema,
        conditions: z.array(conditionSchema).min(1),
        hasHeader: z.boolean().optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const rangeInput = args.range as string;
      const conditions = args.conditions as Array<{
        column: string;
        op: string;
        value: CellValue;
      }>;
      const hasHeader = (args.hasHeader as boolean | undefined) ?? false;

      const wb = readWorkbook(filePath);
      const sheet = wb.sheets.find((s) => s.name === sheetName);
      if (!sheet) {
        const names = wb.sheets.map((s) => s.name).join(", ");
        throw new ToolError(
          "SHEET_NOT_FOUND",
          `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
        );
      }

      const addr = parseRange(rangeInput);
      const rangeValues: CellValue[][] = [];
      for (let r = addr.start.row; r <= addr.end.row; r++) {
        const row: CellValue[] = [];
        for (let c = addr.start.col; c <= addr.end.col; c++) {
          const rowIdx = r - 1;
          const colIdx = c - 1;
          if (
            rowIdx >= 0 &&
            rowIdx < sheet.values.length &&
            colIdx >= 0 &&
            colIdx < sheet.values[rowIdx].length
          ) {
            row.push(sheet.values[rowIdx][colIdx]);
          } else {
            row.push(null);
          }
        }
        rangeValues.push(row);
      }

      const headerRow = hasHeader ? rangeValues[0] : undefined;
      const dataRows = hasHeader ? rangeValues.slice(1) : rangeValues;

      const filtered = dataRows.filter((row) => {
        return conditions.every((cond) => {
          const absCol = resolveColumn(cond.column, addr, headerRow);
          const colIdx = absCol - addr.start.col;
          return matchesCondition(row[colIdx], cond.op, cond.value);
        });
      });

      const result = hasHeader && rangeValues.length > 0
        ? [rangeValues[0], ...filtered]
        : filtered;

      return { values: result };
    })
  );

  // 对应第 2.12 节 (3)：dedupe_range —— 去重
  server.registerTool(
    "dedupe_range",
    {
      title: "区域去重",
      description:
        "删除区域内 keyColumns 完全相同的重复行，只保留第一次出现的行（会修改文件）。被删的行变为空白行，区域外的内容不会被移动。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        range: rangeSchema,
        keyColumns: z.array(z.string().min(1)).min(1),
        hasHeader: z.boolean().optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const rangeInput = args.range as string;
      const keyColumns = args.keyColumns as string[];
      const hasHeader = (args.hasHeader as boolean | undefined) ?? false;

      const addr = parseRange(rangeInput);
      let removedCount = 0;

      await editWorkbook(filePath, (wb) => {
        // CSV 仅单工作表，使用第一张工作表
        const ws =
          path.extname(filePath).toLowerCase() === ".csv"
            ? wb.worksheets[0]
            : wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }

        // 对应第 2.12 节：同 sort 读出 raw 与归一化值
        const rawRows: any[] = [];
        const normRows: CellValue[][] = [];
        for (let r = addr.start.row; r <= addr.end.row; r++) {
          const rawRow: any[] = [];
          const normRow: CellValue[] = [];
          for (let c = addr.start.col; c <= addr.end.col; c++) {
            const cell = ws.getCell(r, c);
            rawRow.push(cell.value);
            normRow.push(normalizeCellValue(cell.value));
          }
          rawRows.push(rawRow);
          normRows.push(normRow);
        }

        const headerRow = hasHeader ? normRows[0] : undefined;
        const dataStart = hasHeader ? 1 : 0;
        const dataRaw = rawRows.slice(dataStart);
        const dataNorm = normRows.slice(dataStart);

        const seen = new Set<string>();
        const keptRaw: any[] = [];

        for (let i = 0; i < dataRaw.length; i++) {
          const keyParts = keyColumns.map((token) => {
            const absCol = resolveColumn(token, addr, headerRow);
            const colIdx = absCol - addr.start.col;
            return JSON.stringify(dataNorm[i][colIdx]);
          });
          const key = keyParts.join("|");
          if (seen.has(key)) {
            removedCount++;
          } else {
            seen.add(key);
            keptRaw.push(dataRaw[i]);
          }
        }

        // 对应第 2.12 节：保留行紧凑写回区域前部，尾部多出来的行整行写 null
        for (let i = 0; i < rawRows.length; i++) {
          if (hasHeader && i === 0) {
            for (let j = 0; j < rawRows[0].length; j++) {
              ws.getCell(addr.start.row + i, addr.start.col + j).value = rawRows[0][j] as any;
            }
          } else {
            const dataIdx = i - dataStart;
            const srcRow = dataIdx < keptRaw.length ? keptRaw[dataIdx] : null;
            for (let j = 0; j < rawRows[0].length; j++) {
              ws.getCell(addr.start.row + i, addr.start.col + j).value = srcRow ? (srcRow[j] as any) : null;
            }
          }
        }
      });

      return { removedCount };
    })
  );
}

/**
 * 对应第 2.12 节：sort_range 的值比较规则
 * - null 永远排最后
 * - 两个都是 number 比数值
 * - 其余情况 String(a).localeCompare(String(b), "zh-Hans-CN")
 * - desc 时取反（null 仍然最后）
 */
function compareValues(
  a: CellValue,
  b: CellValue,
  order: "asc" | "desc"
): number {
  if (a === null && b === null) return 0;
  if (a === null) return 1;
  if (b === null) return -1;

  let result: number;
  if (typeof a === "number" && typeof b === "number") {
    result = a - b;
  } else {
    result = String(a).localeCompare(String(b), "zh-Hans-CN");
  }

  return order === "desc" ? -result : result;
}

/**
 * 对应第 2.12 节：filter_range 条件匹配
 * = / != 使用宽松相等；> / >= / < / <= 按数值或字符串序比较；contains 做子串判断
 */
function matchesCondition(
  cellValue: CellValue,
  op: string,
  value: CellValue
): boolean {
  switch (op) {
    case "=":
      return looseEqual(cellValue, value);
    case "!=":
      return !looseEqual(cellValue, value);
    case ">":
      return compareForRelOp(cellValue, value) > 0;
    case ">=":
      return compareForRelOp(cellValue, value) >= 0;
    case "<":
      return compareForRelOp(cellValue, value) < 0;
    case "<=":
      return compareForRelOp(cellValue, value) <= 0;
    case "contains":
      return String(cellValue).includes(String(value));
    default:
      throw new ToolError("INVALID_PARAMS", `不支持的操作符：${op}`);
  }
}

/** 对应第 2.12 节：= / != 的宽松相等 */
function looseEqual(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  return String(a) === String(b);
}

/**
 * 对应第 2.12 节：> / >= / < / <= 的比较
 * 两侧都不为 null 且都能转为有限 number 时比数值，否则比字符串序
 */
function compareForRelOp(a: CellValue, b: CellValue): number {
  if (a !== null && b !== null) {
    const numA = Number(a);
    const numB = Number(b);
    if (Number.isFinite(numA) && Number.isFinite(numB)) {
      return numA - numB;
    }
  }
  return String(a).localeCompare(String(b), "zh-Hans-CN");
}
