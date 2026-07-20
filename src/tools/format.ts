// 对应 docs/03-detailed-design.md 第 2.11 节：src/tools/format.ts 格式工具

import * as path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { filePathSchema, rangeSchema, tool } from "./common.js";
import { editWorkbook } from "../engine/writer.js";
import { letterToCol, parseRange } from "../engine/address.js";
import { ToolError } from "../engine/errors.js";

/**
 * 对应第 2.11 节：
 * 导出 registerFormatTools(server: McpServer): void
 * 注册 3 个格式工具：format_cells / merge_cells / set_dimensions
 */
export function registerFormatTools(server: McpServer): void {
  // 对应第 2.11 节 (1)：format_cells 的 style schema
  const styleSchema = z
    .object({
      font: z
        .object({
          name: z.string().optional(), // 字体名，如 "微软雅黑"
          size: z.number().positive().optional(),
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          color: z.string().optional(), // 6 位 hex，如 "FF0000"，可带 # 前缀
        })
        .optional(),
      fill: z
        .object({
          color: z.string().optional(), // 纯色填充，6 位 hex
        })
        .optional(),
      border: z
        .object({
          style: z.enum(["thin", "medium", "thick"]).optional(),
          color: z.string().optional(),
        })
        .optional(), // 四边统一设置
      alignment: z
        .object({
          horizontal: z.enum(["left", "center", "right"]).optional(),
          vertical: z.enum(["top", "middle", "bottom"]).optional(),
          wrapText: z.boolean().optional(),
        })
        .optional(),
      numberFormat: z.string().optional(), // 如 "0.00"、"yyyy-mm-dd"、"#,##0"
    })
    .refine(
      (o) => Object.values(o).some((v) => v !== undefined),
      { message: "style 至少要包含一项设置" }
    );

  // 对应第 2.11 节：format_cells —— 设置样式
  server.registerTool(
    "format_cells",
    {
      title: "设置单元格样式",
      description:
        "设置一个区域的字体、填充色、边框、对齐、数字格式，让表格好看。只传想改的项，没传的不动。csv 文件不支持样式。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        range: rangeSchema,
        style: styleSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件不支持样式设置。"
        );
      }
      const sheetName = args.sheetName as string;
      const rangeInput = args.range as string;
      const style = args.style as Record<string, unknown>;
      if (!Object.values(style).some((v) => v !== undefined)) {
        throw new ToolError("INVALID_PARAMS", "style 至少要包含一项设置");
      }
      const addr = parseRange(rangeInput);

      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }

        for (let r = addr.start.row; r <= addr.end.row; r++) {
          for (let c = addr.start.col; c <= addr.end.col; c++) {
            const cell = ws.getCell(r, c);
            const cellStyle = style as {
              font?: Record<string, unknown>;
              fill?: { color?: string };
              border?: { style?: "thin" | "medium" | "thick"; color?: string };
              alignment?: Record<string, unknown>;
              numberFormat?: string;
            };

            // 对应第 2.11 节：font → cell.font
            if (cellStyle.font) {
              const font: Record<string, unknown> = {};
              if (cellStyle.font.name !== undefined) font.name = cellStyle.font.name;
              if (cellStyle.font.size !== undefined) font.size = cellStyle.font.size;
              if (cellStyle.font.bold !== undefined) font.bold = cellStyle.font.bold;
              if (cellStyle.font.italic !== undefined) font.italic = cellStyle.font.italic;
              if (cellStyle.font.color !== undefined) {
                font.color = { argb: toArgb(cellStyle.font.color as string) };
              }
              cell.font = font;
            }

            // 对应第 2.11 节：fill → cell.fill
            if (cellStyle.fill?.color) {
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: toArgb(cellStyle.fill.color) },
              };
            }

            // 对应第 2.11 节：border → cell.border（四边同一对象）
            if (cellStyle.border) {
              const part = {
                style: cellStyle.border.style,
                color: cellStyle.border.color
                  ? { argb: toArgb(cellStyle.border.color) }
                  : undefined,
              };
              cell.border = { top: part, left: part, bottom: part, right: part };
            }

            // 对应第 2.11 节：alignment → cell.alignment
            if (cellStyle.alignment) {
              const alignment: Record<string, unknown> = {};
              if (cellStyle.alignment.horizontal !== undefined) {
                alignment.horizontal = cellStyle.alignment.horizontal;
              }
              if (cellStyle.alignment.vertical !== undefined) {
                alignment.vertical = cellStyle.alignment.vertical;
              }
              if (cellStyle.alignment.wrapText !== undefined) {
                alignment.wrapText = cellStyle.alignment.wrapText;
              }
              cell.alignment = alignment;
            }

            // 对应第 2.11 节：numberFormat → cell.numFmt
            if (cellStyle.numberFormat) {
              cell.numFmt = cellStyle.numberFormat;
            }
          }
        }
      });

      return {};
    })
  );

  // 对应第 2.11 节 (2)：merge_cells —— 合并单元格
  server.registerTool(
    "merge_cells",
    {
      title: "合并单元格",
      description:
        "把一个矩形区域合并成一个单元格，内容保留左上角的值。常用于标题行。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        range: rangeSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件不支持合并单元格。"
        );
      }
      const sheetName = args.sheetName as string;
      const rangeInput = args.range as string;
      const addr = parseRange(rangeInput);

      // 对应第 2.11 节：range 必须跨越多于 1 个单元格，否则 INVALID_PARAMS
      if (addr.start.row === addr.end.row && addr.start.col === addr.end.col) {
        throw new ToolError(
          "INVALID_PARAMS",
          "合并区域必须包含至少两个单元格。"
        );
      }

      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }
        // 对应第 2.11 节：ExcelJS 此 API 为 1 起始，直接用解析结果
        ws.mergeCells(addr.start.row, addr.start.col, addr.end.row, addr.end.col);
      });

      return {};
    })
  );

  // 对应第 2.11 节 (3)：set_dimensions —— 设置列宽行高
  server.registerTool(
    "set_dimensions",
    {
      title: "设置列宽行高",
      description:
        "设置工作表的列宽和/或行高，让表格排版合适。列宽单位是字符数（Excel 默认约 8.43），行高单位是磅（默认约 15）。columns 和 rows 至少传一个。csv 不支持。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        columns: z
          .array(
            z.object({
              column: z
                .string()
                .regex(/^[A-Za-z]+$/)
                .describe("列字母，如 `B`，不区分大小写"),
              width: z.number().positive().describe("列宽，单位字符数，如 16"),
            })
          )
          .min(1)
          .optional(),
        rows: z
          .array(
            z.object({
              row: z.number().int().positive().describe("行号，1 起始"),
              height: z.number().positive().describe("行高，单位磅，如 24"),
            })
          )
          .min(1)
          .optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件不支持设置列宽行高。"
        );
      }
      const sheetName = args.sheetName as string;
      const columns = args.columns as
        | Array<{ column: string; width: number }>
        | undefined;
      const rows = args.rows as
        | Array<{ row: number; height: number }>
        | undefined;

      // 对应第 2.11 节：columns / rows 至少传一个
      if (!columns && !rows) {
        throw new ToolError(
          "INVALID_PARAMS",
          "columns 和 rows 至少要传一个。"
        );
      }

      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }
        // 对应第 2.11 节：列宽 → getColumn(列号).width；行高 → getRow(行号).height
        for (const c of columns ?? []) {
          ws.getColumn(letterToCol(c.column)).width = c.width;
        }
        for (const r of rows ?? []) {
          ws.getRow(r.row).height = r.height;
        }
      });

      return {};
    })
  );
}

/**
 * 对应第 2.11 节：把 6 位 hex 颜色转换为 ExcelJS 的 ARGB 格式
 * 去掉 # 前缀，转大写，前面补 FF alpha
 */
function toArgb(hex: string): string {
  const cleaned = hex.replace(/^#/, "").toUpperCase();
  return `FF${cleaned}`;
}
