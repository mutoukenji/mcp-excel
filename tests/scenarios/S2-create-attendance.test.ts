/**
 * S2 新建 12 张考勤表（需求场景二 / 架构流程三）
 *
 * 涉及模块：tools/workbook.ts、tools/write.ts、engine/writer.ts、engine/reader.ts
 * 步骤：
 *   1. create_workbook(filePath, sheets=["一月",…,"十二月"])；
 *   2. 对每张表各调一次 write_range(startCell="A1", values=[["姓名",1,2,…,31]])；
 *   3. list_sheets + 抽查 3 张表 read_range("A1:AF1") 验证。
 * 预期：sheetsCreated: 12；12 次写入全部成功；每张表首行恰 32 列且内容正确；各次调用互不干扰。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerReadTools } from "../../src/tools/read.js";

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

const MONTHS = [
  "一月","二月","三月","四月","五月","六月",
  "七月","八月","九月","十月","十一月","十二月",
];

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s2-"));
}

describe("S2 新建 12 张考勤表", () => {
  test("S2-01 create_workbook 创建 12 张工作表", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "考勤表.xlsx");

    const res = parseText(
      await handlers.create_workbook({ filePath, sheets: MONTHS })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.sheetsCreated, 12);

    // 验证 list_sheets
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.strictEqual(list.success, true);
    assert.strictEqual(list.sheets.length, 12);
    const names = list.sheets.map((s: any) => s.name);
    for (const m of MONTHS) {
      assert.ok(names.includes(m), `应包含工作表"${m}"`);
    }

    // 所有表初始都是空表
    for (const s of list.sheets) {
      assert.strictEqual(s.rowCount, 0);
      assert.strictEqual(s.colCount, 0);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S2-02 对每张表写入表头行（姓名 + 1~31号）", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "考勤表.xlsx");

    // 创建 12 张表
    await handlers.create_workbook({ filePath, sheets: MONTHS });

    // 准备表头行：["姓名", 1, 2, ..., 31]
    const headerRow: (string | number)[] = ["姓名"];
    for (let d = 1; d <= 31; d++) {
      headerRow.push(d);
    }
    const headerValues = [headerRow];

    // 对每张表写表头
    const writeResults: any[] = [];
    for (const sheetName of MONTHS) {
      const r = parseText(
        await handlers.write_range({
          filePath,
          sheetName,
          startCell: "A1",
          values: headerValues,
        })
      );
      writeResults.push(r);
    }

    // 全部成功
    for (let i = 0; i < writeResults.length; i++) {
      assert.strictEqual(writeResults[i].success, true, `工作表 ${MONTHS[i]} 写入应成功`);
      assert.strictEqual(writeResults[i].cellsWritten, 32, `工作表 ${MONTHS[i]} 应写入 32 格`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S2-03 抽查 3 张表 read_range('A1:AF1') 验证首行内容", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "考勤表.xlsx");

    await handlers.create_workbook({ filePath, sheets: MONTHS });

    const headerRow: (string | number)[] = ["姓名"];
    for (let d = 1; d <= 31; d++) headerRow.push(d);

    for (const sheetName of MONTHS) {
      await handlers.write_range({
        filePath,
        sheetName,
        startCell: "A1",
        values: [headerRow],
      });
    }

    // 抽查 3 张表：一月、六月、十二月
    const spotChecks = ["一月", "六月", "十二月"];
    for (const sheetName of spotChecks) {
      const res = parseText(
        await handlers.read_range({
          filePath,
          sheetName,
          range: "A1:AF1",
        })
      );
      assert.strictEqual(res.success, true, `读"${sheetName}"应成功`);
      assert.strictEqual(res.values.length, 1);
      assert.strictEqual(res.values[0].length, 32, `"${sheetName}"首行应有 32 列`);
      assert.strictEqual(res.values[0][0], "姓名");
      assert.strictEqual(res.values[0][1], 1);
      assert.strictEqual(res.values[0][31], 31);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S2-04 各次调用互不干扰", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "考勤表.xlsx");

    await handlers.create_workbook({ filePath, sheets: MONTHS });

    const headerRow: (string | number)[] = ["姓名"];
    for (let d = 1; d <= 31; d++) headerRow.push(d);

    // 用不同的 startCell 写入不同数据来验证互不干扰
    // 在每张表的 A1 写表头，A2 写"全勤"
    for (const sheetName of MONTHS) {
      await handlers.write_range({
        filePath, sheetName, startCell: "A1", values: [headerRow],
      });
      await handlers.write_range({
        filePath, sheetName, startCell: "A2", values: [["全勤"]],
      });
    }

    // 验证一月、六月、十二月都有正确的首行和 A2="全勤"
    const spotChecks = ["一月", "六月", "十二月"];
    for (const sheetName of spotChecks) {
      const res = parseText(
        await handlers.read_range({ filePath, sheetName, range: "A1:AF2" })
      );
      assert.strictEqual(res.values.length, 2);
      assert.strictEqual(res.values[0].length, 32);
      assert.strictEqual(res.values[1][0], "全勤");
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S2-05 list_sheets 验证最终各表行列数", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "考勤表.xlsx");

    await handlers.create_workbook({ filePath, sheets: MONTHS });

    const headerRow: (string | number)[] = ["姓名"];
    for (let d = 1; d <= 31; d++) headerRow.push(d);

    for (const sheetName of MONTHS) {
      await handlers.write_range({
        filePath, sheetName, startCell: "A1", values: [headerRow],
      });
    }

    const list = parseText(await handlers.list_sheets({ filePath }));
    for (const s of list.sheets) {
      assert.strictEqual(s.rowCount, 1, `${s.name} 应有 1 行数据`);
      assert.strictEqual(s.colCount, 32, `${s.name} 应有 32 列`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });
});
