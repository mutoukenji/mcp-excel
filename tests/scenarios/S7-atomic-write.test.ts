/**
 * S7 原子写故障注入（写坏保护链路）
 *
 * 涉及模块：engine/writer.ts、engine/errors.ts、tools/common.ts
 * 步骤：
 *   1. 对有内容的 xlsx 用文件锁模拟 EPERM → 调用 write_range；
 *   2. 对不存在目录场景 → 调用 create_workbook；
 *   3. 检查原文件字节、临时文件残留、返回包络。
 * 预期：两种故障下原文件字节均不变；无 .tmp-* 残留；分别返回 FILE_BUSY / DIR_NOT_FOUND 人话包络且 isError: true。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWorkbookTools(mockServer as any);
registerWriteTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s7-"));
}

describe("S7 原子写故障注入", () => {
  test("S7-01 文件被占用 → FILE_BUSY 人话包络、isError=true、原文件不变", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "busy.xlsx");

    // 创建文件并写入数据
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath, sheetName: "Sheet1", startCell: "A1",
      values: [["原始数据"]],
    });

    // 记录原文件字节
    const originalBytes = fs.readFileSync(filePath);

    // 用文件锁模拟占用
    const fd = fs.openSync(filePath, "r+");
    try {
      const res = parseText(
        await handlers.write_range({
          filePath,
          sheetName: "Sheet1",
          startCell: "A1",
          values: [["新数据"]],
        })
      );
      assert.strictEqual(res.success, false, "文件被占用应返回失败");

      // 验证 isError（通过检查 fail() 包络的结构）
      // tool() 通过 fail() 返回 isError: true
      const rawRes = await handlers.write_range({
        filePath, sheetName: "Sheet1", startCell: "A1", values: [["新数据"]],
      });
      assert.strictEqual(rawRes.isError, true);

      // 人话文案
      assert.ok(
        res.error.includes("占用") || res.error.includes("Excel"),
        `应提示文件占用，实际: ${res.error}`
      );

      // 原文件字节不变
      const currentBytes = fs.readFileSync(filePath);
      assert.deepStrictEqual(currentBytes, originalBytes, "原文件不应被修改");
    } finally {
      fs.closeSync(fd);
    }

    // 无 .tmp-* 残留
    const tmpFiles = fs.readdirSync(dir).filter((f) => f.includes(".tmp-"));
    assert.strictEqual(tmpFiles.length, 0, `不应有临时文件残留: ${tmpFiles}`);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S7-02 目录不存在 → DIR_NOT_FOUND 人话包络、isError=true", async () => {
    const nonExistentPath = path.join(os.tmpdir(), "mcp-excel-nonexistent-" + Date.now(), "test.xlsx");

    const res = parseText(
      await handlers.create_workbook({ filePath: nonExistentPath, sheets: ["Sheet1"] })
    );
    assert.strictEqual(res.success, false);
    assert.ok(
      res.error.includes("目录") || res.error.includes("不存在"),
      `应提示目录不存在，实际: ${res.error}`
    );

    const rawRes = await handlers.create_workbook({
      filePath: nonExistentPath, sheets: ["Sheet1"],
    });
    assert.strictEqual(rawRes.isError, true);
  });

  test("S7-03 无 .tmp-* 残留（正常失败路径）", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "noresidue.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath, sheetName: "Sheet1", startCell: "A1",
      values: [["数据"]],
    });

    const originalBytes = fs.readFileSync(filePath);

    // 用文件锁制造失败
    const fd = fs.openSync(filePath, "r+");
    try {
      await handlers.write_range({
        filePath, sheetName: "Sheet1", startCell: "A1", values: [["改"]],
      });
    } finally {
      fs.closeSync(fd);
    }

    // 验证原文件未变
    const afterBytes = fs.readFileSync(filePath);
    assert.deepStrictEqual(afterBytes, originalBytes);

    // 验证无临时文件
    const tmpFiles = fs.readdirSync(dir).filter((f) => f.includes(".tmp-"));
    assert.strictEqual(tmpFiles.length, 0);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S7-04 损坏文件被编辑 → 原文件不变", async () => {
    const dir = await tmpDir();
    const corruptPath = path.join(dir, "corrupt.xlsx");
    fs.writeFileSync(corruptPath, Buffer.from([0x50, 0x4b, 0x03, 0x04, 0xff, 0xff, 0xff, 0xff]));

    const originalBytes = fs.readFileSync(corruptPath);

    const res = parseText(
      await handlers.write_range({
        filePath: corruptPath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [[1]],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(
      res.error.includes("无法解析") || res.error.includes("损坏"),
      `应提示文件损坏，实际: ${res.error}`
    );

    // 原文件不变
    const afterBytes = fs.readFileSync(corruptPath);
    assert.deepStrictEqual(afterBytes, originalBytes, "损坏文件不应被修改");

    await fs.promises.rm(dir, { recursive: true });
  });
});
