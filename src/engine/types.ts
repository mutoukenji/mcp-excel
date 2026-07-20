// 对应 docs/03-detailed-design.md 第 2.14 节：src/engine/types.ts 统一内存结构类型

/**
 * 对应第 2.14 节：
 * 单元格值的统一形态，永远是 JSON 可序列化的四种。
 * 日期由 reader.ts 转为 ISO 字符串后也归入 string。
 */
export type CellValue = string | number | boolean | null;

/**
 * 对应第 2.14 节：
 * 文件格式由扩展名判定，仅三种。
 */
export type BookFormat = "xlsx" | "xls" | "csv";

/**
 * 对应第 2.14 节：
 * 统一的工作表内存结构。
 * - rowCount = values.length，已裁掉尾部全空行。
 * - colCount = 所有行中的最大列数，空表为 0。
 * - values 每行长度已补齐到 colCount。
 */
export interface UnifiedSheet {
  name: string;
  rowCount: number;
  colCount: number;
  values: CellValue[][];
}

/**
 * 对应第 2.14 节：
 * 统一的工作簿内存结构，工具层与引擎层之间的唯一数据形态。
 */
export interface UnifiedWorkbook {
  format: BookFormat;
  sheets: UnifiedSheet[];
}
