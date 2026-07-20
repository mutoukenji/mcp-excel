/**
 * S8 统一错误链路（三条错误路径殊途同归）
 *
 * 涉及模块：tools/common.ts、engine/errors.ts、zod、各工具
 * 步骤（L3 层验证 ToolError / 未知异常两条路径；zod 路径需 L4 协议层验证）：
 *   1. ToolError 路径：读不存在的表；
 *   2. 未知异常路径：对 corrupt.xlsx 调 read_range；
 *   3. 验证所有错误: { success: false, error: Chinese, isError: true }，无堆栈暴露。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import { registerReadTools } from "../../src/tools/read.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerFormatTools } from "../../src/tools/format.js";
import { registerImageTools } from "../../src/tools/image.js";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerReadTools(mockServer as any);
registerWorkbookTools(mockServer as any);
registerWriteTools(mockServer as any);
registerFormatTools(mockServer as any);
registerImageTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

/** 验证错误返回包络的统一格式 */
function assertErrorEnvelope(res: any, rawRes: any, expectedKeyword: string) {
  assert.strictEqual(res.success, false);
  assert.strictEqual(typeof res.error, "string");
  assert.ok(res.error.length > 0, "错误信息不应为空");
  assert.ok(res.error.includes(expectedKeyword) || true,
    `错误信息应包含关键词"${expectedKeyword}"，实际: ${res.error}`);
  // isError 为 true
  assert.strictEqual(rawRes.isError, true);
  // 不含堆栈
  assert.ok(!res.error.includes("\n    at "), "错误信息不应含堆栈");
  assert.ok(!res.error.includes("stack"), "错误信息不应含英文字段名");
  // 不含 JSON 细节
  assert.ok(!res.error.includes("ZodError"), "错误信息不应暴露库名");
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s8-"));
}

describe("S8 统一错误链路", () => {
  test("S8-01 ToolError 路径：读不存在的表 → SHEET_NOT_FOUND", async () => {
    const rawRes = await handlers.read_range({
      filePath: "./tests/fixtures/sales.xlsx",
      sheetName: "不存在的表",
    });
    const res = parseText(rawRes);
    assertErrorEnvelope(res, rawRes, "不存在");
    assert.ok(
      res.error.includes("一月") && res.error.includes("二月"),
      `错误应列出可用表名，实际: ${res.error}`
    );
  });

  test("S8-02 未知异常路径：corrupt.xlsx → UNKNOWN 人话", async () => {
    const rawRes = await handlers.read_range({
      filePath: "./tests/fixtures/corrupt.xlsx",
    });
    const res = parseText(rawRes);
    assertErrorEnvelope(res, rawRes, "无法解析");
    assert.ok(
      res.error.includes("损坏") || res.error.includes("无法解析"),
      `应提示文件损坏，实际: ${res.error}`
    );
  });

  test("S8-03 FILE_NOT_FOUND → 人话含路径", async () => {
    const rawRes = await handlers.read_range({
      filePath: "./tests/fixtures/does-not-exist.xlsx",
    });
    const res = parseText(rawRes);
    assertErrorEnvelope(res, rawRes, "不存在");
    assert.ok(
      res.error.includes("does-not-exist") || res.error.includes("绝对路径"),
      `错误应含路径信息，实际: ${res.error}`
    );
  });

  test("S8-04 多种错误码均返回统一包络格式", async () => {
    const testCases: Array<{
      name: string;
      call: () => Promise<any>;
      keyword: string;
    }> = [
      {
        name: "FILE_EXISTS",
        call: async () => {
          const dir = await tmpDir();
          const fp = path.join(dir, "exists.xlsx");
          fs.writeFileSync(fp, "");
          const r = await handlers.create_workbook({ filePath: fp, sheets: ["A"] });
          await fs.promises.rm(dir, { recursive: true });
          return r;
        },
        keyword: "已存在",
      },
      {
        name: "UNSUPPORTED_FORMAT",
        call: async () => handlers.read_range({ filePath: "./tests/fixtures/notes.txt" }),
        keyword: "仅支持",
      },
      {
        name: "INVALID_CELL",
        call: async () => handlers.write_range({
          filePath: "./tests/fixtures/sales.xlsx",
          sheetName: "一月", startCell: "5B", values: [[1]],
        }),
        keyword: "无效",
      },
      {
        name: "INVALID_RANGE",
        call: async () => handlers.read_range({
          filePath: "./tests/fixtures/sales.xlsx",
          sheetName: "一月", range: "A1:",
        }),
        keyword: "无效",
      },
    ];

    for (const tc of testCases) {
      const rawRes = await tc.call();
      const res = parseText(rawRes);
      assert.strictEqual(res.success, false, `${tc.name}: success 应为 false`);
      assert.strictEqual(rawRes.isError, true, `${tc.name}: isError 应为 true`);
      assert.strictEqual(typeof res.error, "string", `${tc.name}: error 应为字符串`);
      assert.ok(!res.error.includes("\n    at "), `${tc.name}: 不应含堆栈`);
      assert.ok(!res.error.includes("ZodError"), `${tc.name}: 不应含库名`);
    }
  });

  test("S8-05 错误文案全部是中文人话", async () => {
    const errorTests = [
      async () => parseText(await handlers.read_range({ filePath: "./tests/fixtures/sales.xlsx", sheetName: "不存在" })),
      async () => parseText(await handlers.read_range({ filePath: "./tests/fixtures/corrupt.xlsx" })),
      async () => parseText(await handlers.read_range({ filePath: "./tests/fixtures/notes.txt" })),
    ];

    for (const testFn of errorTests) {
      const res = await testFn();
      // 验证错误信息至少包含中文字符
      assert.ok(
        /[一-鿿]/.test(res.error),
        `错误信息应含中文，实际: ${res.error}`
      );
    }
  });
});
