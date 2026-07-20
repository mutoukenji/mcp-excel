// 对应 docs/03-detailed-design.md 第 2.8 节：src/tools/workbook.ts 工作簿/工作表管理工具

import * as path from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { filePathSchema, sheetNameSchema, tool } from "./common.js";
import { readWorkbook } from "../engine/reader.js";
import { createWorkbook, editWorkbook } from "../engine/writer.js";
import { ToolError } from "../engine/errors.js";

/**
 * 对应第 2.8 节：
 * 导出 registerWorkbookTools(server: McpServer): void
 * 注册 5 个工作簿/工作表管理工具：
 * create_workbook / list_sheets / add_sheet / delete_sheet / rename_sheet
 */
export function registerWorkbookTools(server: McpServer): void {
  // 对应第 2.8 节 (1)：create_workbook —— 新建 Excel 文件
  server.registerTool(
    "create_workbook",
    {
      title: "新建 Excel 文件",
      description:
        "从无到有新建一个 .xlsx 文件。可指定初始工作表名列表。文件已存在时会失败（防误覆盖），不会清空已有文件。",
      inputSchema: {
        filePath: filePathSchema,
        sheets: z.array(sheetNameSchema).min(1).optional(),
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const ext = path.extname(filePath).toLowerCase();
      if (ext !== ".xlsx") {
        throw new ToolError("INVALID_PARAMS", "新建文件的扩展名必须是 .xlsx");
      }
      const sheets = (args.sheets as string[] | undefined) ?? ["Sheet1"];
      const set = new Set(sheets);
      if (set.size !== sheets.length) {
        throw new ToolError("INVALID_PARAMS", "初始工作表名称存在重复");
      }
      await createWorkbook(filePath, sheets);
      return { sheetsCreated: sheets.length };
    })
  );

  // 对应第 2.8 节 (2)：list_sheets —— 列出所有工作表
  server.registerTool(
    "list_sheets",
    {
      title: "列出工作表",
      description:
        "查看一个 Excel 文件里有哪些工作表，各有多少行多少列。动手读/写之前可以先调它确认表名。",
      inputSchema: {
        filePath: filePathSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      const wb = readWorkbook(filePath);
      const sheets = wb.sheets.map((s) => ({
        name: s.name,
        rowCount: s.rowCount,
        colCount: s.colCount,
      }));
      return { sheets };
    })
  );

  // 对应第 2.8 节 (3)：add_sheet —— 新增一张工作表
  server.registerTool(
    "add_sheet",
    {
      title: "新增工作表",
      description: "在已有文件中追加一张空工作表。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: sheetNameSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件只有一张工作表，不支持新增"
        );
      }
      const sheetName = args.sheetName as string;
      await editWorkbook(filePath, (wb) => {
        if (wb.getWorksheet(sheetName)) {
          throw new ToolError("SHEET_EXISTS", `工作表 "${sheetName}" 已存在。`);
        }
        wb.addWorksheet(sheetName);
      });
      return {};
    })
  );

  // 对应第 2.8 节 (4)：delete_sheet —— 删除一张工作表
  server.registerTool(
    "delete_sheet",
    {
      title: "删除工作表",
      description: "删除指定工作表。不能删除最后一张。",
      inputSchema: {
        filePath: filePathSchema,
        sheetName: sheetNameSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件只有一张工作表，不支持删除"
        );
      }
      const sheetName = args.sheetName as string;
      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${sheetName}" 不存在。现有工作表：${names}。`
          );
        }
        if (wb.worksheets.length === 1) {
          throw new ToolError("LAST_SHEET", "不能删除最后一张工作表。");
        }
        wb.removeWorksheet(ws.id);
      });
      return {};
    })
  );

  // 对应第 2.8 节 (5)：rename_sheet —— 重命名工作表
  server.registerTool(
    "rename_sheet",
    {
      title: "重命名工作表",
      description: "修改工作表名称。",
      inputSchema: {
        filePath: filePathSchema,
        oldName: z.string().min(1),
        newName: sheetNameSchema,
      },
    },
    tool(async (args) => {
      const filePath = args.filePath as string;
      if (path.extname(filePath).toLowerCase() === ".csv") {
        throw new ToolError(
          "UNSUPPORTED_FORMAT",
          "CSV 文件只有一张工作表，不支持重命名"
        );
      }
      const oldName = args.oldName as string;
      const newName = args.newName as string;
      await editWorkbook(filePath, (wb) => {
        const ws = wb.getWorksheet(oldName);
        if (!ws) {
          const names = wb.worksheets.map((w) => w.name).join(", ");
          throw new ToolError(
            "SHEET_NOT_FOUND",
            `工作表 "${oldName}" 不存在。现有工作表：${names}。`
          );
        }
        const existing = wb.getWorksheet(newName);
        if (existing && newName !== oldName) {
          throw new ToolError("SHEET_EXISTS", `工作表 "${newName}" 已存在。`);
        }
        ws.name = newName;
      });
      return {};
    })
  );
}
