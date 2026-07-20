/**
 * S6 格式支持矩阵全量核对（全局约定 8）
 *
 * 涉及模块：全部工具 × reader/writer 的格式分支
 * 对同一份内容的 .xlsx / .xls / .csv 三个副本，分别执行代表性操作。
 * 预期每个（操作 × 格式）组合结果与设计 0.8 节矩阵逐格一致：
 *   xlsx 全 ✅；
 *   xls 读 ✅、写一律 READ_ONLY_FORMAT；
 *   csv 读写值 ✅（sort/dedupe 可用）、表操作与样式类一律 UNSUPPORTED_FORMAT 且文案为 CSV 专用文案。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerReadTools } from "../../src/tools/read.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerDataTools } from "../../src/tools/data.js";
import { registerFormatTools } from "../../src/tools/format.js";
import { registerImageTools } from "../../src/tools/image.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWorkbookTools(mockServer as any);
registerReadTools(mockServer as any);
registerWriteTools(mockServer as any);
registerDataTools(mockServer as any);
registerFormatTools(mockServer as any);
registerImageTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s6-"));
}

const fixturesDir = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "fixtures"
);

// CSV 的 3 个读操作与 4 个写操作
const READ_OPS = ["list_sheets", "read_range", "filter_range"] as const;
const WRITE_OPS = ["write_range", "set_formula", "sort_range", "dedupe_range"] as const;
const SHEET_OPS = ["add_sheet", "delete_sheet", "rename_sheet"] as const;
const STYLE_OPS = ["format_cells", "merge_cells", "insert_image", "set_dimensions"] as const;

describe("S6 格式支持矩阵", () => {
  // ===== xlsx 全功能 =====
  test("S6-01 xlsx 全功能可用", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "test.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath, sheetName: "Sheet1", startCell: "A1",
      values: [["姓名", "销售额"], ["张三", 1000]],
    });

    // 读操作
    assert.strictEqual(parseText(await handlers.list_sheets({ filePath })).success, true);
    assert.strictEqual(
      parseText(await handlers.read_range({ filePath, sheetName: "Sheet1" })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.filter_range({
        filePath, sheetName: "Sheet1", range: "A1:B2", hasHeader: true,
        conditions: [{ column: "销售额", op: ">", value: 500 }],
      })).success, true
    );

    // 写操作
    assert.strictEqual(
      parseText(await handlers.write_range({
        filePath, sheetName: "Sheet1", startCell: "B2", values: [[2000]],
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.set_formula({
        filePath, sheetName: "Sheet1", cell: "B3", formula: "SUM(B2)",
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.sort_range({
        filePath, sheetName: "Sheet1", range: "A1:B2", keyColumn: "B", order: "asc",
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.dedupe_range({
        filePath, sheetName: "Sheet1", range: "A1:B2", keyColumns: ["A"],
      })).success, true
    );

    // 表操作
    assert.strictEqual(
      parseText(await handlers.add_sheet({ filePath, sheetName: "新表" })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.rename_sheet({ filePath, oldName: "新表", newName: "改名" })).success, true
    );
    const delRes = parseText(
      await handlers.delete_sheet({ filePath, sheetName: "改名" })
    );
    assert.strictEqual(delRes.success, true);

    // 样式操作
    assert.strictEqual(
      parseText(await handlers.format_cells({
        filePath, sheetName: "Sheet1", range: "A1:A1",
        style: { font: { bold: true } },
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.merge_cells({
        filePath, sheetName: "Sheet1", range: "A1:A2",
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.insert_image({
        filePath, sheetName: "Sheet1",
        imagePath: path.join(fixturesDir, "img.png"),
        anchorCell: "D2", width: 80, height: 80,
      })).success, true
    );
    assert.strictEqual(
      parseText(await handlers.set_dimensions({
        filePath, sheetName: "Sheet1", columns: [{ column: "A", width: 20 }],
      })).success, true
    );

    await fs.promises.rm(dir, { recursive: true });
  });

  // ===== xls 只读 =====
  test("S6-02 xls 读操作全部可用", async () => {
    const r1 = parseText(await handlers.list_sheets({ filePath: "./tests/fixtures/legacy.xls" }));
    assert.strictEqual(r1.success, true);
    const r2 = parseText(await handlers.read_range({ filePath: "./tests/fixtures/legacy.xls", sheetName: "一月" }));
    assert.strictEqual(r2.success, true);
    assert.strictEqual(r2.values[0][0], "姓名");
    const r3 = parseText(await handlers.filter_range({
      filePath: "./tests/fixtures/legacy.xls", sheetName: "一月", range: "A1:D6",
      conditions: [{ column: "B", op: ">", value: 1000 }], hasHeader: true,
    }));
    assert.strictEqual(r3.success, true);
  });

  test("S6-03 xls 写操作报 READ_ONLY_FORMAT", async () => {
    const xlsPath = "./tests/fixtures/legacy.xls";
    // write_range
    let r = parseText(await handlers.write_range({ filePath: xlsPath, sheetName: "一月", startCell: "A1", values: [[1]] }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"), `应提示 READ_ONLY_FORMAT，实际：${r.error}`);
    // set_formula
    r = parseText(await handlers.set_formula({ filePath: xlsPath, sheetName: "一月", cell: "A1", formula: "SUM(A1)" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    // sort_range
    r = parseText(await handlers.sort_range({ filePath: xlsPath, sheetName: "一月", range: "A1:D6", keyColumn: "B", order: "asc" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    // dedupe_range
    r = parseText(await handlers.dedupe_range({ filePath: xlsPath, sheetName: "一月", range: "A1:D6", keyColumns: ["A"] }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
  });

  test("S6-04 xls 表操作报 READ_ONLY_FORMAT", async () => {
    const xlsPath = "./tests/fixtures/legacy.xls";
    let r = parseText(await handlers.add_sheet({ filePath: xlsPath, sheetName: "新表" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    r = parseText(await handlers.delete_sheet({ filePath: xlsPath, sheetName: "一月" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    r = parseText(await handlers.rename_sheet({ filePath: xlsPath, oldName: "一月", newName: "1月" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
  });

  test("S6-05 xls 样式操作报 READ_ONLY_FORMAT", async () => {
    const xlsPath = "./tests/fixtures/legacy.xls";
    let r = parseText(await handlers.format_cells({ filePath: xlsPath, sheetName: "一月", range: "A1:A1", style: { font: { bold: true } } }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    r = parseText(await handlers.merge_cells({ filePath: xlsPath, sheetName: "一月", range: "A1:A2" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    r = parseText(await handlers.insert_image({ filePath: xlsPath, sheetName: "一月", imagePath: path.join(fixturesDir, "img.png"), anchorCell: "D2" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
    r = parseText(await handlers.set_dimensions({ filePath: xlsPath, sheetName: "一月", columns: [{ column: "A", width: 20 }] }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("老版 .xls"));
  });

  // ===== csv 支持矩阵 =====
  test("S6-06 csv 读操作全部可用", async () => {
    const r1 = parseText(await handlers.list_sheets({ filePath: "./tests/fixtures/data.csv" }));
    assert.strictEqual(r1.success, true);
    const r2 = parseText(await handlers.read_range({ filePath: "./tests/fixtures/data.csv" }));
    assert.strictEqual(r2.success, true);
    const r3 = parseText(await handlers.filter_range({
      filePath: "./tests/fixtures/data.csv", sheetName: "Sheet1", range: "A1:C3",
      conditions: [{ column: "B", op: ">", value: 500 }], hasHeader: true,
    }));
    assert.strictEqual(r3.success, true);
  });

  test("S6-07 csv 读写值操作可用（write_range/sort/dedupe）", async () => {
    const dir = await tmpDir();
    const csvPath = path.join(dir, "test.csv");
    fs.writeFileSync(csvPath, "a,b\n1,2\n3,4", "utf-8");

    // write_range
    const w = parseText(await handlers.write_range({ filePath: csvPath, sheetName: "ignored", startCell: "A3", values: [[5, 6]] }));
    assert.strictEqual(w.success, true, `write_range csv: ${w.error || ""}`);

    // sort_range
    const s = parseText(await handlers.sort_range({ filePath: csvPath, sheetName: "ignored", range: "A1:B3", keyColumn: "A", order: "desc" }));
    assert.strictEqual(s.success, true, `sort csv: ${s.error || ""}`);

    // dedupe_range
    const d = parseText(await handlers.dedupe_range({ filePath: csvPath, sheetName: "ignored", range: "A1:B3", keyColumns: ["A"] }));
    assert.strictEqual(d.success, true, `dedupe csv: ${d.error || ""}`);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S6-08 csv 表操作报 UNSUPPORTED_FORMAT（CSV 专用文案）", async () => {
    const csvPath = "./tests/fixtures/data.csv";
    let r = parseText(await handlers.add_sheet({ filePath: csvPath, sheetName: "新表" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"), `CSV 表操作应含 CSV 文案，实际: ${r.error}`);
    assert.ok(!r.error.includes("老版 .xls"), "不应混杂 xls 文案");

    r = parseText(await handlers.delete_sheet({ filePath: csvPath, sheetName: "Sheet1" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"));
    assert.ok(!r.error.includes("老版 .xls"));

    r = parseText(await handlers.rename_sheet({ filePath: csvPath, oldName: "Sheet1", newName: "新名" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"));
    assert.ok(!r.error.includes("老版 .xls"));
  });

  test("S6-09 csv 样式操作报 UNSUPPORTED_FORMAT（CSV 专用文案）", async () => {
    const csvPath = "./tests/fixtures/data.csv";
    let r = parseText(await handlers.format_cells({ filePath: csvPath, sheetName: "Sheet1", range: "A1:A1", style: { font: { bold: true } } }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"), `CSV 样式操作应含 CSV 文案，实际: ${r.error}`);
    assert.ok(!r.error.includes("老版 .xls"));

    r = parseText(await handlers.merge_cells({ filePath: csvPath, sheetName: "Sheet1", range: "A1:A2" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"));
    assert.ok(!r.error.includes("老版 .xls"));

    r = parseText(await handlers.insert_image({ filePath: csvPath, sheetName: "Sheet1", imagePath: path.join(fixturesDir, "img.png"), anchorCell: "D2" }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"));
    assert.ok(!r.error.includes("老版 .xls"));

    r = parseText(await handlers.set_dimensions({ filePath: csvPath, sheetName: "Sheet1", columns: [{ column: "A", width: 20 }] }));
    assert.strictEqual(r.success, false);
    assert.ok(r.error.includes("CSV"));
    assert.ok(!r.error.includes("老版 .xls"));
  });

  test("S6-10 csv set_formula 可用", async () => {
    const dir = await tmpDir();
    const csvPath = path.join(dir, "test.csv");
    fs.writeFileSync(csvPath, "a,b\n1,2\n3,4", "utf-8");
    const r = parseText(await handlers.set_formula({ filePath: csvPath, sheetName: "ignored", cell: "C1", formula: "SUM(A1:B1)" }));
    assert.strictEqual(r.success, true);
    await fs.promises.rm(dir, { recursive: true });
  });
});
