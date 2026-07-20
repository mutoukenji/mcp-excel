/**
 * S9 MCP 端到端冒烟（模块 1 + 全量注册）
 *
 * 涉及模块：src/index.ts、tools/index.ts、全部工具
 * 步骤：用 MCP Client 经 stdio 启动真实进程 → initialize → tools/list →
 *   tools/call 依次调 create_workbook、write_range、read_range → 关闭进程。
 * 预期：tools/list 返回恰好 15 个工具；create → write → read 完整链路成功；
 *   stdout 只出现 JSON-RPC 消息；进程可正常退出。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const rootDir = path.resolve(fileURLToPath(import.meta.url), "..", "..", "..");
const distIndex = path.join(rootDir, "dist", "index.js");

function parseToolResult(result: any) {
  return JSON.parse((result.content[0] as { text: string }).text);
}

describe("S9 MCP 端到端冒烟", () => {
  test("S9-01 tools/list 返回 15 个工具且各有中文描述", async () => {
    const transport = new StdioClientTransport({
      command: "node",
      args: [distIndex],
      cwd: rootDir,
      stderr: "pipe",
    });

    let stderr = "";
    const errStream = transport.stderr;
    if (errStream) {
      errStream.setEncoding("utf8");
      errStream.on("data", (d: string) => (stderr += d));
    }

    const client = new Client({ name: "test-s9", version: "1.0.0" });
    await client.connect(transport);

    try {
      // verify tools/list
      const tools = await client.listTools();
      assert.strictEqual(tools.tools.length, 15,
        `应注册 15 个工具，实际 ${tools.tools.length}`);

      for (const t of tools.tools) {
        assert.ok(t.title, `工具 ${t.name} 缺少 title`);
        assert.ok(t.description, `工具 ${t.name} 缺少 description`);
        const text = (t.title ?? "") + (t.description ?? "");
        assert.ok(
          /[一-龥]/.test(text),
          `工具 ${t.name} 的 title/description 缺少中文`
        );
      }

      const toolNames = tools.tools.map((t) => t.name);
      const expected = [
        "create_workbook", "list_sheets", "add_sheet", "delete_sheet", "rename_sheet",
        "read_range",
        "write_range", "set_formula",
        "format_cells", "merge_cells", "set_dimensions",
        "sort_range", "filter_range", "dedupe_range",
        "insert_image",
      ];
      for (const name of expected) {
        assert.ok(toolNames.includes(name), `缺少工具: ${name}`);
      }

      // verify stderr has startup message
      assert.ok(stderr.includes("已启动") || stderr.includes("mcp-excel"),
        `stderr 应有启动日志，实际: ${stderr}`);
    } finally {
      await client.close();
    }
  });

  test("S9-02 create_workbook → write_range → read_range 完整工具链", async () => {
    const transport = new StdioClientTransport({
      command: "node",
      args: [distIndex],
      cwd: rootDir,
      stderr: "pipe",
    });

    let stderr = "";
    const errStream = transport.stderr;
    if (errStream) {
      errStream.setEncoding("utf8");
      errStream.on("data", (d: string) => (stderr += d));
    }

    const client = new Client({ name: "test-s9-chain", version: "1.0.0" });
    await client.connect(transport);

    const tmpDir = await fs.promises.mkdtemp(
      path.join(os.tmpdir(), "mcp-excel-s9-")
    );

    try {
      const filePath = path.join(tmpDir, "e2e-chain.xlsx");

      // create_workbook
      const createRes = parseToolResult(
        await client.callTool({
          name: "create_workbook",
          arguments: { filePath, sheets: ["数据表"] },
        })
      );
      assert.strictEqual(createRes.success, true);
      assert.strictEqual(createRes.sheetsCreated, 1);

      // write_range
      const writeRes = parseToolResult(
        await client.callTool({
          name: "write_range",
          arguments: {
            filePath,
            sheetName: "数据表",
            startCell: "A1",
            values: [["项目","金额"],["收入",5000],["支出",2000]],
          },
        })
      );
      assert.strictEqual(writeRes.success, true);
      assert.strictEqual(writeRes.cellsWritten, 6);

      // read_range
      const readRes = parseToolResult(
        await client.callTool({
          name: "read_range",
          arguments: { filePath, sheetName: "数据表" },
        })
      );
      assert.strictEqual(readRes.success, true);
      assert.strictEqual(readRes.values[0][0], "项目");
      assert.strictEqual(readRes.values[1][1], 5000);

      // list_sheets roundtrip
      const listRes = parseToolResult(
        await client.callTool({
          name: "list_sheets",
          arguments: { filePath },
        })
      );
      assert.strictEqual(listRes.success, true);
      assert.strictEqual(listRes.sheets[0].name, "数据表");
      assert.strictEqual(listRes.sheets[0].rowCount, 3);
    } finally {
      await client.close();
      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    }
  });
});
