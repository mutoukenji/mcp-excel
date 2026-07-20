// 对应 docs/03-detailed-design.md 第 2.10 节：src/tools/write.ts 写入工具

import * as path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import {
  filePathSchema,
  cellSchema,
  valuesSchema,
  tool,
} from "./common.js";
import { editWorkbook } from "../engine/writer.js";
import { parseCell } from "../engine/address.js";
import { ToolError } from "../engine/errors.js";
import type { CellValue } from "../engine/types.js";

/**
 * 对应第 2.10 节：
 * 导出 registerWriteTools(server: McpServer): void
 * 注册 2 个写入工具：write_range / set_formula
 */
export function registerWriteTools(server: McpServer): void {
  // 对应第 2.10 节 (1)：write_range —— 写入一块数据
  server.registerTool(
    "write_range",
    {
      title: "写入区域数据",
      description:
        "从 startCell 开始写入一个二维数组（先行后列）。一个格子就传 `[[值]]`；一大片就传多行多列。值为 null 会清空对应单元格。字符串不会被套公式——要写公式请用 set_formula。目标工作表必须已存在（没有就先用 add_sheet）。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        startCell: cellSchema,
        values: valuesSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const startCell = args.startCell as string;
      const values = args.values as CellValue[][];

      const { row, col } = parseCell(startCell);
      let cellsWritten = 0;

      await editWorkbook(filePath, (wb) => {
        // 对应第 2.10 节：csv 文件忽略 sheetName，直接写唯一工作表
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

        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            // 对应第 2.10 节：v === null 赋 null 即清空；ExcelJS 行/列均为 1 起始
            ws.getCell(row + i, col + j).value = values[i][j] as any;
            cellsWritten++;
          }
        }
      });

      return { cellsWritten };
    })
  );

  // 对应第 2.10 节 (2)：set_formula —— 在单元格写入公式
  server.registerTool(
    "set_formula",
    {
      title: "写入公式",
      description:
        "在指定单元格写入 Excel 公式，如 `SUM(A1:A10)`，可带或不带开头的 `=`。注意：公式要等用户在 Excel/WPS 里打开文件后才会计算出结果，本工具读该格会得到 null。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        cell: cellSchema,
        formula: z.string().min(1).describe("公式文本，如 `SUM(A1:A10)` 或 `=AVERAGE(B2:B30)`"),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const cell = args.cell as string;
      let formula = (args.formula as string).replace(/^=/, "");

      // 对应第 2.10 节：去掉开头 '=' 后为空串 → INVALID_PARAMS
      if (!formula) {
        throw new ToolError("INVALID_PARAMS", "公式内容不能为空");
      }

      const { row, col } = parseCell(cell);

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

        ws.getCell(row, col).value = { formula };
      });

      return {};
    })
  );
}
