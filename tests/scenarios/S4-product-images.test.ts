/**
 * S4 商品清单插图（需求场景四）
 *
 * 涉及模块：tools/workbook.ts、tools/write.ts、tools/image.ts、engine/writer.ts
 * 步骤：
 *   1. create_workbook 新建"商品清单.xlsx"（含"清单"表）；
 *   2. write_range 写入 3 行商品数据；
 *   3. 对每个商品 insert_image(anchorCell="D2"/"D3"/"D4", width=120, height=120) 插 3 张图；
 *   4. ExcelJS 读回验证。
 * 预期：3 张图分别锚定 D2/D3/D4、尺寸正确；单元格数据不受影响。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerImageTools } from "../../src/tools/image.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWorkbookTools(mockServer as any);
registerWriteTools(mockServer as any);
registerImageTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s4-"));
}

const fixturesDir = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "fixtures"
);

describe("S4 商品清单插图", () => {
  test("S4-01 创建商品清单并写入数据", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "商品清单.xlsx");

    // Step 1: 创建
    const createRes = parseText(
      await handlers.create_workbook({ filePath, sheets: ["清单"] })
    );
    assert.strictEqual(createRes.success, true);
    assert.strictEqual(createRes.sheetsCreated, 1);

    // Step 2: 写入 3 行商品数据
    const writeRes = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "清单",
        startCell: "A1",
        values: [
          ["商品名称", "价格", "库存", "图片"],
          ["商品A", 99.9, 100, null],
          ["商品B", 199.0, 50, null],
          ["商品C", 299.0, 30, null],
        ],
      })
    );
    assert.strictEqual(writeRes.success, true);
    assert.strictEqual(writeRes.cellsWritten, 16);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S4-02 插入 3 张图片到对应商品行", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "商品清单.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["清单"] });
    await handlers.write_range({
      filePath,
      sheetName: "清单",
      startCell: "A1",
      values: [
        ["商品名称", "价格", "库存", "图片"],
        ["商品A", 99.9, 100, null],
        ["商品B", 199.0, 50, null],
        ["商品C", 299.0, 30, null],
      ],
    });

    // Step 3: 对每个商品插入图片
    const imgPath = path.join(fixturesDir, "img.png");
    const anchors = ["D2", "D3", "D4"];
    for (const anchor of anchors) {
      const res = parseText(
        await handlers.insert_image({
          filePath,
          sheetName: "清单",
          imagePath: imgPath,
          anchorCell: anchor,
          width: 120,
          height: 120,
        })
      );
      assert.strictEqual(res.success, true, `插图到 ${anchor} 应成功`);
    }

    // Step 4: ExcelJS 读回验证
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("清单")!;

    const images = ws.getImages();
    assert.strictEqual(images.length, 3, "应有 3 张图片");

    // 验证每张图的锚点与尺寸
    const anchorsExpected = [
      { col: 3, row: 1 },  // D2: col=4 → 0-based=3, row=2 → 0-based=1
      { col: 3, row: 2 },  // D3
      { col: 3, row: 3 },  // D4
    ];
    for (let i = 0; i < 3; i++) {
      const img = images[i];
      assert.strictEqual(img.range.tl.col, anchorsExpected[i].col,
        `第 ${i + 1} 张图片列锚点应为 ${anchorsExpected[i].col}`);
      assert.strictEqual(img.range.tl.row, anchorsExpected[i].row,
        `第 ${i + 1} 张图片行锚点应为 ${anchorsExpected[i].row}`);
      assert.strictEqual(img.range.ext.width, 120,
        `第 ${i + 1} 张图片宽度应为 120`);
      assert.strictEqual(img.range.ext.height, 120,
        `第 ${i + 1} 张图片高度应为 120`);
    }

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S4-03 图片不影响单元格数据", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "商品清单2.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["清单"] });
    await handlers.write_range({
      filePath,
      sheetName: "清单",
      startCell: "A1",
      values: [
        ["商品名称", "价格"],
        ["商品A", 99.9],
      ],
    });
    await handlers.insert_image({
      filePath,
      sheetName: "清单",
      imagePath: path.join(fixturesDir, "img.png"),
      anchorCell: "C2",
      width: 100,
      height: 100,
    });

    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "商品名称");
    assert.strictEqual(wb.sheets[0].values[0][1], "价格");
    assert.strictEqual(wb.sheets[0].values[1][0], "商品A");
    assert.strictEqual(wb.sheets[0].values[1][1], 99.9);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S4-04 不同图片格式插入", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "商品清单3.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["清单"] });
    await handlers.write_range({
      filePath,
      sheetName: "清单",
      startCell: "A1",
      values: [["测试"]],
    });

    // png, jpg, gif 三种格式各插一张
    const formats = [
      { ext: "png", path: path.join(fixturesDir, "img.png") },
      { ext: "jpg", path: path.join(fixturesDir, "img.jpg") },
      { ext: "gif", path: path.join(fixturesDir, "img.gif") },
    ];

    for (let i = 0; i < formats.length; i++) {
      const res = parseText(
        await handlers.insert_image({
          filePath,
          sheetName: "清单",
          imagePath: formats[i].path,
          anchorCell: `D${i + 1}`,
          width: 80,
          height: 80,
        })
      );
      assert.strictEqual(res.success, true,
        `插入 ${formats[i].ext} 格式图片应成功`);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("清单")!;
    assert.strictEqual(ws.getImages().length, 3,
      "应有 3 张不同格式的图片");

    await fs.promises.rm(dir, { recursive: true });
  });
});
