import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import ExcelJS from "exceljs";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWriteTools(mockServer as any);
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-write-"));
}

describe("src/tools/write.ts", () => {
  test("WR-01 单格写入", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "B5",
        values: [[3500]],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.cellsWritten, 1);
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[4][1], 3500);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-02 批量写入", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [
          [1, 2, 3, 4],
          [5, 6, 7, 8],
        ],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.cellsWritten, 8);
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], 1);
    assert.strictEqual(wb.sheets[0].values[1][3], 8);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-03 各行长度不一致", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [[1, 2, 3], [4]],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.cellsWritten, 4);
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values[0], [1, 2, 3]);
    assert.deepStrictEqual(wb.sheets[0].values[1], [4, null, null]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-04 null 清空单元格", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["x", "y"]],
    });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [[null]],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.cellsWritten, 1);
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], null);
    assert.strictEqual(wb.sheets[0].values[0][1], "y");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-05 =... 字符串不当公式", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w5.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["=1+1"]],
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "=1+1");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-06 表不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w6.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "不存在的表",
        startCell: "A1",
        values: [[1]],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "不存在的表" 不存在'));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-07 非法 startCell", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "w7.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A0",
        values: [[1]],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("无效的单元格地址"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-08 对 xls 写入", async () => {
    const before = fs.readFileSync("./tests/fixtures/legacy.xls");
    const res = parseText(
      await handlers.write_range({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        startCell: "A1",
        values: [[1]],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("老版 .xls"));
    const after = fs.readFileSync("./tests/fixtures/legacy.xls");
    assert.deepStrictEqual(before, after);
  });

  test("WR-09 csv 忽略 sheetName", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "data.csv");
    fs.copyFileSync("./tests/fixtures/data.csv", filePath);
    const res = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "任意名",
        startCell: "A2",
        values: [["测试"]],
      })
    );
    assert.strictEqual(res.success, true);
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[1][0], "测试");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-10 文件不存在", async () => {
    const res = parseText(
      await handlers.write_range({
        filePath: "./tests/fixtures/not-exist.xlsx",
        sheetName: "Sheet1",
        startCell: "A1",
        values: [[1]],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("文件不存在"));
  });

  test("WR-11 文件被占用", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "busy.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const fd = fs.openSync(filePath, "r+");
    try {
      const res = parseText(
        await handlers.write_range({
          filePath,
          sheetName: "Sheet1",
          startCell: "A1",
          values: [[1]],
        })
      );
      assert.strictEqual(res.success, false);
      assert.ok(res.error.includes("占用"));
    } finally {
      fs.closeSync(fd);
    }
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-12 set_formula 不带等号", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "formula.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.set_formula({
        filePath,
        sheetName: "Sheet1",
        cell: "B10",
        formula: "SUM(B2:B9)",
      })
    );
    assert.strictEqual(res.success, true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("B10");
    assert.strictEqual((cell.value as any).formula, "SUM(B2:B9)");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-13 set_formula 带等号", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "formula2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.set_formula({
      filePath,
      sheetName: "Sheet1",
      cell: "B10",
      formula: "=AVERAGE(B2:B9)",
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("B10");
    assert.strictEqual((cell.value as any).formula, "AVERAGE(B2:B9)");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-14 公式去 = 后为空", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "formula3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.set_formula({
        filePath,
        sheetName: "Sheet1",
        cell: "B10",
        formula: "=",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("公式内容不能为空"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-15 set_formula 已知限制", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "formula4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        [1, 1],
        [1, null],
      ],
    });
    await handlers.set_formula({
      filePath,
      sheetName: "Sheet1",
      cell: "B2",
      formula: "A1+1",
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[1][1], null);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WR-16 set_formula 对 xls", async () => {
    const res = parseText(
      await handlers.set_formula({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        cell: "B10",
        formula: "SUM(B2:B9)",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("老版 .xls"));
  });
});
