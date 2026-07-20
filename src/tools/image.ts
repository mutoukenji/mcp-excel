// 对应 docs/03-detailed-design.md 第 2.13 节：src/tools/image.ts 图片工具

import * as fs from "node:fs";
import * as path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { filePathSchema, cellSchema, tool } from "./common.js";
import { editWorkbook } from "../engine/writer.js";
import { parseCell } from "../engine/address.js";
import { ToolError } from "../engine/errors.js";

/**
 * 对应第 2.13 节：
 * 导出 registerImageTools(server: McpServer): void
 * 注册 1 个图片工具：insert_image
 */
export function registerImageTools(server: McpServer): void {
  server.registerTool(
    "insert_image",
    {
      title: "插入图片",
      description:
        "把一张本地图片（png / jpg / gif）插入工作表，图片左上角锚定在 anchorCell。width/height 为像素，不传给默认 300×200——比例不对时请显式传这两个值。csv 不支持插图。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: z.string().min(1),
        imagePath: z.string().min(1),
        anchorCell: cellSchema,
        width: z.number().positive().optional(),
        height: z.number().positive().optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const sheetName = args.sheetName as string;
      const imagePath = args.imagePath as string;
      const anchorCell = args.anchorCell as string;
      const width = args.width as number | undefined;
      const height = args.height as number | undefined;

      // 对应第 2.13 节：csv 不支持插图
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError("UNSUPPORTED_FORMAT", "CSV 文件不支持插入图片。");
      }

      // 对应第 2.13 节：图片不存在 → IMAGE_NOT_FOUND
      if (!fs.existsSync(imagePath)) {
        throw new ToolError(
          "IMAGE_NOT_FOUND",
          `图片文件不存在：${imagePath}。请检查路径。`
        );
      }

      // 对应第 2.13 节：扩展名映射
      const ext = path.extname(imagePath).toLowerCase();
      let extension: "png" | "jpeg" | "gif";
      switch (ext) {
        case ".png":
          extension = "png";
          break;
        case ".jpg":
        case ".jpeg":
          extension = "jpeg";
          break;
        case ".gif":
          extension = "gif";
          break;
        default:
          throw new ToolError(
            "UNSUPPORTED_IMAGE",
            `不支持的图片格式：${ext}，仅支持 png / jpg / gif。`
          );
      }

      const { row, col } = parseCell(anchorCell);

      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }

        // 对应第 2.13 节：添加图片并锚定到单元格
        const imageId = wb.addImage({ filename: imagePath, extension });
        ws.addImage(imageId, {
          tl: { col: col - 1, row: row - 1 }, // ExcelJS 的 tl 是 0 起始
          ext: { width: width ?? 300, height: height ?? 200 },
        });
      });

      return {};
    })
  );
}
