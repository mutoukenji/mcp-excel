#!/usr/bin/env node
// 对应 docs/03-detailed-design.md 第 2.5 节：src/index.ts 进程入口

// 对应第 2.5 节：从 MCP SDK 引入服务器与 stdio 传输层
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAllTools } from "./tools/index.js";

/**
 * 对应第 2.5 节表格：
 * - 函数 main() 无输入，输出 Promise<void>
 * - 职责：创建 server → 注册全部工具 → 连接 stdio
 */
async function main(): Promise<void> {
  const server = new McpServer({
    // 对应第 2.5 节：McpServer 的 name / version 固定取值
    name: "mcp-excel",
    version: "0.1.0",
  });

  // 对应第 2.5 节：注册全部 14 个工具（由 src/tools/index.ts 汇总）
  registerAllTools(server);

  // 对应第 2.5 节：建立 stdio 传输通道并连接
  const transport = new StdioServerTransport();
  await server.connect(transport);

  // 对应第 2.5 节：仅允许通过 stderr 输出日志（全局约定 3：stdout 是协议通道）
  console.error("mcp-excel 已启动，通过 stdio 等待 AI 助手连接");
}

// 对应第 2.5 节：启动失败时输出到 stderr 并退出进程，错误码 1
main().catch((e: unknown) => {
  console.error("mcp-excel 启动失败：", e);
  process.exit(1);
});
