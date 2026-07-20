// 对应 docs/03-detailed-design.md 第 2.9 节：src/tools/read.ts 读取工具

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { filePathSchema, rangeSchema, tool } from "./common.js";
import { readWorkbook } from "../engine/reader.js";
import { parseRange } from "../engine/address.js";
import { ToolError } from "../engine/errors.js";
import type { CellValue } from "../engine/types.js";

/**
 * 对应第 2.9 节：
 * 导出 registerReadTools(server: McpServer): void
 * 注册 1 个读取工具：read_range
 */
export function registerReadTools(server: McpServer): void {
  server.registerTool(
    "read_range",
    {
      title: "读取区域内容",
      description:
        "读取一张工作表的内容，返回二维数组。不传 range 则读整张表；表很大时建议先传 range 读一部分（如前 50 行）了解结构。空单元格为 null，日期为 ISO 字符串。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1).optional(),
        range: rangeSchema.optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const wb = readWorkbook(filePath);
      const sheetName = args.sheetName as string | undefined;

      // 对应第 2.9 节：sheetName 未传取 sheets[0]；传了但找不到 → SHEET_NOT_FOUND（附现有表名）
      const sheet = sheetName
        ? wb.sheets.find((s) => s.name === sheetName)
        : wb.sheets[0];
      if (!sheet) {
        const names = wb.sheets.map((s) => s.name).join(", ");
        throw new ToolError(
          "SHEET_NOT_FOUND",
          `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
        );
      }

      const rangeInput = args.range as string | undefined;

      // 对应第 2.9 节：未传 range 直接返回整张表
      if (!rangeInput) {
        return { values: sheet.values };
      }

      // 对应第 2.9 节：传了 range → parseRange 解析，矩形内超出实际数据范围补 null
      const addr = parseRange(rangeInput);
      const rowCount = addr.end.row - addr.start.row + 1;
      const colCount = addr.end.col - addr.start.col + 1;
      const values: CellValue[][] = [];

      for (let i = 0; i < rowCount; i++) {
        const rowIndex = addr.start.row + i - 1; // 转为 0 起始索引
        const row: CellValue[] = [];
        for (let j = 0; j < colCount; j++) {
          const colIndex = addr.start.col + j - 1; // 转为 0 起始索引
          if (
            rowIndex >= 0 &&
            rowIndex < sheet.values.length &&
            colIndex >= 0 &&
            colIndex < sheet.values[rowIndex].length
          ) {
            row.push(sheet.values[rowIndex][colIndex]);
          } else {
            row.push(null);
          }
        }
        values.push(row);
      }

      return { values };
    })
  );
}
