// 对应 docs/03-detailed-design.md 第 2.6 节：src/tools/index.ts 汇总注册入口

import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerWorkbookTools } from "./workbook.js";
import { registerReadTools } from "./read.js";
import { registerWriteTools } from "./write.js";
import { registerFormatTools } from "./format.js";
import { registerDataTools } from "./data.js";
import { registerImageTools } from "./image.js";

/**
 * 对应第 2.6 节表格：
 * - 函数 registerAllTools(server: McpServer): void
 * - 职责：依次调用 6 个分文件的注册函数，把全部 14 个工具挂到 MCP server
 *
 * 第 2.6 节说明：注册失败属于启动失败，由 src/index.ts 的 main().catch() 兜底。
 */
export function registerAllTools(server: McpServer): void {
  // 对应第 2.6 节：工作簿 / 工作表管理工具（5 个）
  registerWorkbookTools(server);

  // 对应第 2.6 节：读取工具（1 个）
  registerReadTools(server);

  // 对应第 2.6 节：写入工具（2 个）
  registerWriteTools(server);

  // 对应第 2.6 节：格式工具（2 个）
  registerFormatTools(server);

  // 对应第 2.6 节：数据工具（3 个）
  registerDataTools(server);

  // 对应第 2.6 节：图片工具（1 个）
  registerImageTools(server);
}
