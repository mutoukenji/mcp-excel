import { test, describe } from "node:test";
import assert from "node:assert";
import { registerReadTools } from "../../src/tools/read.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerReadTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

describe("src/tools/read.ts", () => {
  test("RD-01 读整张表", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );
    assert.strictEqual(res.success, true);
    assert.ok(res.values.length > 0);
  });

  test("RD-02 默认第一张表", async () => {
    const res = parseText(
      await handlers.read_range({ filePath: "./tests/fixtures/sales.xlsx" })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values[0][0], "姓名");
  });

  test("RD-03 指定区域且超出数据范围", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
        range: "A1:B10",
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 10);
    assert.strictEqual(res.values[0].length, 2);
    assert.deepStrictEqual(res.values[9], [null, null]);
  });

  test("RD-04 区域行列数严格等于请求", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
        range: "C2:D4",
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 3);
    assert.strictEqual(res.values[0].length, 2);
  });

  test("RD-05 表名打错", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月份",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "一月份" 不存在'));
    assert.ok(res.error.includes("一月"));
    assert.ok(res.error.includes("二月"));
  });

  test("RD-06 非法区域", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
        range: "A1:",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("无效的区域"));
  });

  test("RD-07 空表", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "二月",
      })
    );
    assert.strictEqual(res.success, true);
    assert.deepStrictEqual(res.values, []);
  });

  test("RD-08 读 xls / csv", async () => {
    const xls = parseText(
      await handlers.read_range({ filePath: "./tests/fixtures/legacy.xls" })
    );
    assert.strictEqual(xls.success, true);
    assert.strictEqual(xls.values[0][0], "姓名");
    const csv = parseText(
      await handlers.read_range({ filePath: "./tests/fixtures/data.csv" })
    );
    assert.strictEqual(csv.success, true);
    assert.strictEqual(csv.values[0][0], "姓名");
  });

  test("RD-09 日期与空单元格表现", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );
    const date = res.values[1][2];
    assert.ok(typeof date === "string" && date.startsWith("2026-01-"));
    assert.strictEqual(res.values[1][3], null);
  });
});
