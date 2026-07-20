/**
 * S10 无状态与连续调用（全局约定 1）
 *
 * 涉及模块：engine/writer.ts 与全部写类工具
 * 步骤：
 *   1. 对同一文件快速连续调用 20 次 write_range（不同单元格），随后 read_range 全量读回；
 *   2. 对两个不同文件交替调用写操作。
 * 预期：20 次写入全部生效、无丢失无串扰；两文件内容各自正确。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerReadTools } from "../../src/tools/read.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWorkbookTools(mockServer as any);
registerWriteTools(mockServer as any);
registerReadTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s10-"));
}

describe("S10 无状态与连续调用", () => {
  test("S10-01 同一文件连续写入 20 次全部生效", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "consecutive.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });

    // 连续 20 次写入（每次写不同单元格）
    const expectedValues: Record<string, number> = {};
    for (let i = 0; i < 20; i++) {
      const row = i + 1;
      const value = i * 100;
      const col = "A";
      const cellAddr = `${col}${row}`;
      expectedValues[`${row - 1}`] = value;

      const res = parseText(
        await handlers.write_range({
          filePath,
          sheetName: "Sheet1",
          startCell: cellAddr,
          values: [[value]],
        })
      );
      assert.strictEqual(res.success, true,
        `第 ${i + 1} 次写入 ${cellAddr} 失败: ${res.error || ""}`);
    }

    // 全量读回验证
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].rowCount, 20,
      `应有 20 行，实际 ${wb.sheets[0].rowCount}`);

    for (let i = 0; i < 20; i++) {
      assert.strictEqual(wb.sheets[0].values[i][0], i * 100,
        `第 ${i + 1} 行值应为 ${i * 100}`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S10-02 两文件交替写入，内容各自正确", async () => {
    const dir = await tmpDir();
    const fileA = path.join(dir, "fileA.xlsx");
    const fileB = path.join(dir, "fileB.xlsx");

    await handlers.create_workbook({ filePath: fileA, sheets: ["Sheet1"] });
    await handlers.create_workbook({ filePath: fileB, sheets: ["Sheet1"] });

    // 交替写入 10 轮（共 20 次写入）
    for (let i = 0; i < 10; i++) {
      // 写 fileA
      const rA = parseText(
        await handlers.write_range({
          filePath: fileA,
          sheetName: "Sheet1",
          startCell: `A${i + 1}`,
          values: [[`A-${i + 1}`]],
        })
      );
      assert.strictEqual(rA.success, true, `fileA 第 ${i + 1} 轮写入失败`);

      // 写 fileB
      const rB = parseText(
        await handlers.write_range({
          filePath: fileB,
          sheetName: "Sheet1",
          startCell: `A${i + 1}`,
          values: [[`B-${i + 1}`]],
        })
      );
      assert.strictEqual(rB.success, true, `fileB 第 ${i + 1} 轮写入失败`);
    }

    // 验证 fileA
    const wbA = readWorkbook(fileA);
    assert.strictEqual(wbA.sheets[0].rowCount, 10);
    for (let i = 0; i < 10; i++) {
      assert.strictEqual(wbA.sheets[0].values[i][0], `A-${i + 1}`);
    }

    // 验证 fileB（应完全独立）
    const wbB = readWorkbook(fileB);
    assert.strictEqual(wbB.sheets[0].rowCount, 10);
    for (let i = 0; i < 10; i++) {
      assert.strictEqual(wbB.sheets[0].values[i][0], `B-${i + 1}`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S10-03 连续写入无数据丢失或串扰", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "integrity.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });

    // 在不同区域快速写入，验证互不干扰
    // 写 A 列：1~5
    for (let i = 1; i <= 5; i++) {
      await handlers.write_range({
        filePath, sheetName: "Sheet1", startCell: `A${i}`, values: [[`colA-${i}`]],
      });
    }
    // 写 B 列：1~5（与 A 列交替）
    for (let i = 1; i <= 5; i++) {
      await handlers.write_range({
        filePath, sheetName: "Sheet1", startCell: `B${i}`, values: [[`colB-${i}`]],
      });
    }

    // 读回验证：每行两列都应正确（两列写入无串扰）
    const res = parseText(
      await handlers.read_range({ filePath, sheetName: "Sheet1" })
    );
    assert.strictEqual(res.success, true);
    // readWorkbook 会对齐补齐
    for (let i = 0; i < 5; i++) {
      assert.strictEqual(res.values[i][0], `colA-${i + 1}`, `A${i + 1} 值不对`);
      assert.strictEqual(res.values[i][1], `colB-${i + 1}`, `B${i + 1} 值不对`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S10-04 写入后立即读取一致性", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "instant.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });

    // 写入，然后立即读取，验证读到的就是刚写的
    const testData = [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ];

    await handlers.write_range({
      filePath, sheetName: "Sheet1", startCell: "A1", values: testData,
    });

    const res = parseText(
      await handlers.read_range({ filePath, sheetName: "Sheet1" })
    );
    assert.strictEqual(res.success, true);
    assert.deepStrictEqual(res.values, testData);

    await fs.promises.rm(dir, { recursive: true });
  });
});
