import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import ExcelJS from "exceljs";
import { registerFormatTools } from "../../src/tools/format.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerFormatTools(mockServer as any);
registerWriteTools(mockServer as any);
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-fmt-"));
}

describe("src/tools/format.ts", () => {
  test("F-01 字体设置", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f1.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.format_cells({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A1",
        style: { font: { bold: true, size: 12, color: "FF0000" } },
      })
    );
    assert.strictEqual(res.success, true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual(cell.font.bold, true);
    assert.strictEqual(cell.font.size, 12);
    assert.strictEqual((cell.font.color as any).argb, "FFFF0000");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-02 颜色带 # 前缀", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { font: { color: "#305496" } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual((cell.font.color as any).argb, "FF305496");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-03 填充色", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { fill: { color: "305496" } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual((cell.fill as any).type, "pattern");
    assert.strictEqual((cell.fill as any).fgColor.argb, "FF305496");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-04 边框四边统一", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { border: { style: "thin", color: "999999" } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    for (const side of ["top", "left", "bottom", "right"]) {
      assert.strictEqual((cell.border as any)[side].style, "thin");
      assert.strictEqual((cell.border as any)[side].color.argb, "FF999999");
    }
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-05 对齐与自动换行", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f5.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { alignment: { horizontal: "center", vertical: "middle", wrapText: true } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual(cell.alignment.horizontal, "center");
    assert.strictEqual(cell.alignment.vertical, "middle");
    assert.strictEqual(cell.alignment.wrapText, true);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-06 数字格式", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f6.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { numberFormat: "0.00" },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual(cell.numFmt, "0.00");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-07 只传部分项", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f7.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { font: { bold: true } },
    });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A1",
      style: { fill: { color: "FF0000" } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const cell = wb.getWorksheet("Sheet1")!.getCell("A1");
    assert.strictEqual(cell.font.bold, true);
    assert.strictEqual((cell.fill as any).fgColor.argb, "FFFF0000");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-08 style 为空对象", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f8.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.format_cells({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A1",
        style: {},
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("style 至少要包含一项设置"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-09 区域内每格都生效", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f9.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.format_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:D1",
      style: { font: { bold: true } },
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    for (let c = 1; c <= 4; c++) {
      assert.strictEqual(ws.getCell(1, c).font.bold, true);
    }
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-10 对 csv / xls", async () => {
    const csv = parseText(
      await handlers.format_cells({
        filePath: "./tests/fixtures/data.csv",
        sheetName: "Sheet1",
        range: "A1:A1",
        style: { font: { bold: true } },
      })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV 文件不支持样式设置"));
    const xls = parseText(
      await handlers.format_cells({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        range: "A1:A1",
        style: { font: { bold: true } },
      })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
  });

  test("F-11 merge_cells 正常", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f11.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["标题"]],
    });
    await handlers.merge_cells({
      filePath,
      sheetName: "Sheet1",
      range: "A1:D1",
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.ok(ws.getCell("A1").isMerged);
    assert.strictEqual(ws.getCell("A1").value, "标题");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-12 单格区域", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f12.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.merge_cells({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("至少两个单元格"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-13 颠倒区域", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f13.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["标题"]],
    });
    const res = parseText(
      await handlers.merge_cells({
        filePath,
        sheetName: "Sheet1",
        range: "D1:A1",
      })
    );
    assert.strictEqual(res.success, true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.ok(ws.getCell("A1").isMerged);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-14 与已有合并区域重叠", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f14.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["标题"]],
    });
    await handlers.merge_cells({ filePath, sheetName: "Sheet1", range: "A1:D1" });
    const res = parseText(
      await handlers.merge_cells({ filePath, sheetName: "Sheet1", range: "B1:E1" })
    );
    assert.strictEqual(res.success, false);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-15 merge_cells 对 csv", async () => {
    const res = parseText(
      await handlers.merge_cells({
        filePath: "./tests/fixtures/data.csv",
        sheetName: "Sheet1",
        range: "A1:B1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("CSV 文件不支持合并单元格"));
  });

  test("F-16 set_dimensions 设置列宽", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f16.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.set_dimensions({
        filePath,
        sheetName: "Sheet1",
        columns: [{ column: "B", width: 16 }],
      })
    );
    assert.strictEqual(res.success, true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.strictEqual(ws.getColumn(2).width, 16);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-17 set_dimensions 设置行高", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f17.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.set_dimensions({
      filePath,
      sheetName: "Sheet1",
      rows: [{ row: 1, height: 24 }],
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.strictEqual(ws.getRow(1).height, 24);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-18 列宽行高同时设置、列字母小写", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f18.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.set_dimensions({
      filePath,
      sheetName: "Sheet1",
      columns: [{ column: "c", width: 20 }],
      rows: [{ row: 2, height: 30 }],
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.strictEqual(ws.getColumn(3).width, 20);
    assert.strictEqual(ws.getRow(2).height, 30);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-19 columns / rows 都未传", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f19.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.set_dimensions({
        filePath,
        sheetName: "Sheet1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("columns 和 rows 至少要传一个"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-20 set_dimensions 表不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f20.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.set_dimensions({
        filePath,
        sheetName: "X",
        columns: [{ column: "B", width: 16 }],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "X" 不存在'));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("F-21 set_dimensions 对 csv / xls", async () => {
    const csv = parseText(
      await handlers.set_dimensions({
        filePath: "./tests/fixtures/data.csv",
        sheetName: "Sheet1",
        columns: [{ column: "B", width: 16 }],
      })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV 文件不支持设置列宽行高"));
    const xls = parseText(
      await handlers.set_dimensions({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        columns: [{ column: "B", width: 16 }],
      })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
  });

  test("F-22 设置尺寸不影响单元格内容", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "f22.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["x", "y"]],
    });
    await handlers.set_dimensions({
      filePath,
      sheetName: "Sheet1",
      columns: [{ column: "B", width: 16 }],
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "x");
    assert.strictEqual(wb.sheets[0].values[0][1], "y");
    await fs.promises.rm(dir, { recursive: true });
  });
});
