import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { registerImageTools } from "../../src/tools/image.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerImageTools(mockServer as any);
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-img-"));
}

const fixturesDir = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "fixtures"
);

describe("src/tools/image.ts", () => {
  test("I-01 插入 png 显式尺寸", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.insert_image({
        filePath,
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "D2",
        width: 120,
        height: 120,
      })
    );
    assert.strictEqual(res.success, true);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    assert.strictEqual(ws.getImages().length, 1);
    const img = ws.getImages()[0];
    assert.strictEqual(img.range.tl.col, 3);
    assert.strictEqual(img.range.tl.row, 1);
    assert.strictEqual(img.range.ext.width, 120);
    assert.strictEqual(img.range.ext.height, 120);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-02 默认尺寸", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.insert_image({
      filePath,
      sheetName: "Sheet1",
      imagePath: path.join(fixturesDir, "img.png"),
      anchorCell: "A1",
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    const img = ws.getImages()[0];
    assert.strictEqual(img.range.ext.width, 300);
    assert.strictEqual(img.range.ext.height, 200);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-03 扩展名映射", async () => {
    for (const ext of ["jpg", "jpeg", "gif"]) {
      const dir = await tmpDir();
      const filePath = path.join(dir, `img.${ext}.xlsx`);
      await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
      const res = parseText(
        await handlers.insert_image({
          filePath,
          sheetName: "Sheet1",
          imagePath: path.join(fixturesDir, `img.${ext}`),
          anchorCell: "A1",
        })
      );
      assert.strictEqual(res.success, true, ext);
      await fs.promises.rm(dir, { recursive: true });
    }
  });

  test("I-04 不支持的图片格式", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img-bmp.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.insert_image({
        filePath,
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.bmp"),
        anchorCell: "A1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("仅支持 png / jpg / gif"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-05 图片不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img-none.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.insert_image({
        filePath,
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "not-exist.png"),
        anchorCell: "A1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("图片文件不存在"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-06 锚点非法", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img-a0.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.insert_image({
        filePath,
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "A0",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("无效的单元格地址"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-07 表不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img-sheet.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.insert_image({
        filePath,
        sheetName: "X",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "A1",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "X" 不存在'));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("I-08 对 csv / xls", async () => {
    const csv = parseText(
      await handlers.insert_image({
        filePath: "./tests/fixtures/data.csv",
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "A1",
      })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV 文件不支持插入图片"));
    const xls = parseText(
      await handlers.insert_image({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "A1",
      })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
  });

  test("I-09 锚点 0 起始换算", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "img-origin.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.insert_image({
      filePath,
      sheetName: "Sheet1",
      imagePath: path.join(fixturesDir, "img.png"),
      anchorCell: "A1",
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;
    const img = ws.getImages()[0];
    assert.strictEqual(img.range.tl.col, 0);
    assert.strictEqual(img.range.tl.row, 0);
    await fs.promises.rm(dir, { recursive: true });
  });
});
