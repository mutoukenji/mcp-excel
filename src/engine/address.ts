// 对应 docs/03-detailed-design.md 第 2.16 节：src/engine/address.ts 地址/列解析

import { ToolError } from "./errors.js";
import type { CellValue } from "./types.js";

// 对应第 2.16 节：单元格地址，row / col 均为 1 起始
export interface CellAddr {
  row: number;
  col: number;
}

// 对应第 2.16 节：矩形区域，start 恒在左上、end 恒在右下
export interface RangeAddr {
  start: CellAddr;
  end: CellAddr;
}

/**
 * 对应第 2.16 节：
 * "B5" → { row: 5, col: 2 }
 * 正则不匹配或行号为 0 → ToolError("INVALID_CELL")
 */
export function parseCell(input: string): CellAddr {
  const match = input.match(/^([A-Za-z]+)([0-9]+)$/);
  if (!match) {
    throw new ToolError(
      "INVALID_CELL",
      `无效的单元格地址："${input}"。正确示例："B5"。`
    );
  }
  const col = letterToCol(match[1]);
  const row = parseInt(match[2], 10);
  if (row === 0) {
    throw new ToolError(
      "INVALID_CELL",
      `无效的单元格地址："${input}"。正确示例："B5"。`
    );
  }
  return { row, col };
}

/**
 * 对应第 2.16 节：
 * "A1:D20" → 两端各 parseCell；若顺序颠倒（如 "D20:A1"）自动交换为左上→右下，不报错。
 * 格式不合法 → ToolError("INVALID_RANGE")
 */
export function parseRange(input: string): RangeAddr {
  const parts = input.split(":");
  if (parts.length !== 2) {
    throw new ToolError(
      "INVALID_RANGE",
      `无效的区域："${input}"。正确示例："A1:D20"。`
    );
  }
  let start: CellAddr;
  let end: CellAddr;
  try {
    start = parseCell(parts[0]);
    end = parseCell(parts[1]);
  } catch (e) {
    if (e instanceof ToolError && e.code === "INVALID_CELL") {
      throw new ToolError(
        "INVALID_RANGE",
        `无效的区域："${input}"。正确示例："A1:D20"。`
      );
    }
    throw e;
  }

  // 保证 start 在左上，end 在右下
  if (
    start.col > end.col ||
    (start.col === end.col && start.row > end.row)
  ) {
    [start, end] = [end, start];
  }

  return { start, end };
}

/**
 * 对应第 2.16 节：
 * { row: 5, col: 2 } → "B5"
 */
export function toCellName(addr: CellAddr): string {
  return `${colToLetter(addr.col)}${addr.row}`;
}

/**
 * 对应第 2.16 节：
 * 1 → "A"，28 → "AB"
 */
export function colToLetter(col: number): string {
  let result = "";
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

/**
 * 对应第 2.16 节：
 * "ab" → 28（不区分大小写）
 */
export function letterToCol(letters: string): number {
  let col = 0;
  for (const ch of letters.toUpperCase()) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return col;
}

/**
 * 对应第 2.16 节：
 * 解析列标识（列字母或表头名），返回工作表绝对列号（1 起始）。
 * - 纯字母：letterToCol，并校验落在 [range.start.col, range.end.col] 内。
 * - 表头名：要求 headerRow 已提供，精确匹配；未命中附全部表头。
 */
export function resolveColumn(
  token: string,
  range: RangeAddr,
  headerRow?: CellValue[]
): number {
  if (/^[A-Za-z]+$/.test(token)) {
    const col = letterToCol(token);
    if (col < range.start.col || col > range.end.col) {
      throw new ToolError(
        "INVALID_PARAMS",
        `列 ${token} 不在区域 ${toCellName(range.start)}:${toCellName(
          range.end
        )} 内。`
      );
    }
    return col;
  }

  if (!headerRow) {
    throw new ToolError(
      "INVALID_PARAMS",
      "使用表头名指定列时，请设置 hasHeader: true。"
    );
  }

  const index = headerRow.findIndex((v) => String(v) === token);
  if (index === -1) {
    const headers = headerRow.map(String).join(", ");
    throw new ToolError(
      "INVALID_PARAMS",
      `列 "${token}" 不存在。表头包含：${headers}。`
    );
  }

  return range.start.col + index;
}
