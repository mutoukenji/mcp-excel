import * as fs from "node:fs";
import { test, describe } from "node:test";
import assert from "node:assert";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { readWorkbook } from "../../src/engine/reader.js";
import { ToolError } from "../../src/engine/errors.js";

const fixturesDir = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "fixtures"
);

describe("src/engine/reader.ts", () => {
  test("R-01 读 xlsx 多表", () => {
    const wb = readWorkbook(path.join(fixturesDir, "sales.xlsx"));
    assert.strictEqual(wb.format, "xlsx");
    assert.strictEqual(wb.sheets.length, 2);
    assert.strictEqual(wb.sheets[0].name, "一月");
    assert.ok(wb.sheets[0].rowCount > 0);
    assert.strictEqual(wb.sheets[1].name, "二月");
  });

  test("R-02 值归一化", () => {
    const wb = readWorkbook(path.join(fixturesDir, "sales.xlsx"));
    const sheet = wb.sheets[0];
    assert.strictEqual(sheet.values[0][0], "姓名");
    assert.strictEqual(typeof sheet.values[1][1], "number");
    assert.strictEqual(sheet.values[1][1], 1200);
    const date = sheet.values[1][2];
    assert.ok(typeof date === "string" && date.startsWith("2026-01-"));
  });

  test("R-03 行补齐与中间空行保留", () => {
    const wb = readWorkbook(path.join(fixturesDir, "sales.xlsx"));
    const sheet = wb.sheets[0];
    const colCount = sheet.colCount;
    for (const row of sheet.values) {
      assert.strictEqual(row.length, colCount);
    }
    // 中间存在空白行
    assert.ok(sheet.values.some((row) => row.every((v) => v === null)));
  });

  test("R-04 尾部全空行裁掉", () => {
    const wb = readWorkbook(path.join(fixturesDir, "sales.xlsx"));
    const sheet = wb.sheets[0];
    const last = sheet.values[sheet.values.length - 1];
    assert.ok(!last.every((v) => v === null));
  });

  test("R-05 空表", () => {
    const wb = readWorkbook(path.join(fixturesDir, "sales.xlsx"));
    const sheet = wb.sheets.find((s) => s.name === "二月")!;
    assert.strictEqual(sheet.rowCount, 0);
    assert.strictEqual(sheet.colCount, 0);
    assert.deepStrictEqual(sheet.values, []);
  });

  test("R-06 读 xls", () => {
    const wb = readWorkbook(path.join(fixturesDir, "legacy.xls"));
    assert.strictEqual(wb.format, "xls");
    const sheet = wb.sheets[0];
    assert.strictEqual(sheet.values[0][0], "姓名");
    assert.strictEqual(sheet.values[1][1], 1200);
  });

  test("R-07 读 csv", () => {
    const wb = readWorkbook(path.join(fixturesDir, "data.csv"));
    assert.strictEqual(wb.format, "csv");
    assert.strictEqual(wb.sheets.length, 1);
    assert.strictEqual(wb.sheets[0].values[0][0], "姓名");
  });

  test("R-08 文件不存在", () => {
    assert.throws(
      () => readWorkbook(path.join(fixturesDir, "not-exist.xlsx")),
      (e) =>
        e instanceof ToolError &&
        e.code === "FILE_NOT_FOUND" &&
        e.message.includes("not-exist.xlsx") &&
        e.message.includes("需要绝对路径")
    );
  });

  test("R-09 不支持的扩展名", () => {
    assert.throws(
      () => readWorkbook(path.join(fixturesDir, "notes.txt")),
      (e) =>
        e instanceof ToolError &&
        e.code === "UNSUPPORTED_FORMAT" &&
        e.message.includes(".xlsx / .xls / .csv")
    );
  });

  test("R-10 文件损坏", () => {
    assert.throws(
      () => readWorkbook(path.join(fixturesDir, "corrupt.xlsx")),
      (e) =>
        e instanceof ToolError &&
        e.code === "UNKNOWN" &&
        e.message.includes("无法解析文件") &&
        e.message.includes("corrupt.xlsx")
    );
  });

  test("R-11 公式格无缓存结果", async () => {
    const ExcelJS = (await import("exceljs")).default;
    const tmp = path.join(fixturesDir, "formula-no-cache.xlsx");
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("B2").value = { formula: "A1+A2" };
    await wb.xlsx.writeFile(tmp);
    const result = readWorkbook(tmp);
    assert.strictEqual(result.sheets[0].values[1][1], null);
    fs.unlinkSync(tmp);
  });

  test("R-12 富文本/错误值等异类", async () => {
    const ExcelJS = (await import("exceljs")).default;
    const tmp = path.join(fixturesDir, "weird.xlsx");
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = { richText: [{ text: "hello" }, { text: " world" }] };
    ws.getCell("B1").value = 1;
    ws.getCell("A2").value = { error: "#VALUE!" };
    ws.getCell("B2").value = 2;
    await wb.xlsx.writeFile(tmp);
    const result = readWorkbook(tmp);
    assert.strictEqual(result.sheets[0].values[0][0], "hello world");
    assert.strictEqual(result.sheets[0].values[1][0], null);
    fs.unlinkSync(tmp);
  });
});
