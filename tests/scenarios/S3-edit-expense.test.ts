/**
 * S3 改表并加合计行（需求场景三 / 架构流程二）
 *
 * 涉及模块：tools/read.ts、tools/write.ts、engine/writer.ts、engine/address.ts
 * 步骤：
 *   1. 准备含"费用表"的 xlsx（B5=3000，共 9 行）；
 *   2. read_range 确认 B5 位置与末行行号；
 *   3. write_range(startCell="B5", values=[[3500]])；
 *   4. write_range 在第 10 行写 ["合计", null]；
 *   5. set_formula(cell="B10", formula="SUM(B2:B9)")；
 *   6. 读回全表验证。
 * 预期：B5=3500；A10="合计"；ExcelJS 验证 B10 公式为 SUM(B2:B9)；其余单元格不变。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import ExcelJS from "exceljs";
import { registerReadTools } from "../../src/tools/read.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerReadTools(mockServer as any);
registerWriteTools(mockServer as any);
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s3-"));
}

describe("S3 改表并加合计行", () => {
  test("S3-01 准备费用表并修改 B5 为 3500", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "费用表.xlsx");

    // Step 1: 创建费用表，填充 9 行数据
    await handlers.create_workbook({ filePath, sheets: ["费用表"] });
    await handlers.write_range({
      filePath,
      sheetName: "费用表",
      startCell: "A1",
      values: [
        ["项目", "金额"],
        ["办公用品", 1500],
        ["差旅费", 2500],
        ["餐饮费", 800],
        ["其他", 3000],
        ["通讯费", 200],
        ["交通费", 400],
        ["资料费", 600],
        ["杂费", 100],
      ],
    });

    // Step 2: read_range 确认 B5 值
    const before = parseText(
      await handlers.read_range({
        filePath,
        sheetName: "费用表",
        range: "A1:B9",
      })
    );
    assert.strictEqual(before.success, true);
    // B5 = row 5, col 2 → values[4][1]
    assert.strictEqual(before.values[4][1], 3000,
      "修改前 B5 应为 3000");
    // 确认末行行号
    assert.strictEqual(before.values.length, 9);

    // Step 3: 修改 B5 为 3500
    const writeRes = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "费用表",
        startCell: "B5",
        values: [[3500]],
      })
    );
    assert.strictEqual(writeRes.success, true);
    assert.strictEqual(writeRes.cellsWritten, 1);

    // Step 4: 在第 10 行写 ["合计", null]
    const sumLabel = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "费用表",
        startCell: "A10",
        values: [["合计", null]],
      })
    );
    assert.strictEqual(sumLabel.success, true);

    // Step 5: 在 B10 写公式 SUM(B2:B9)
    const formulaRes = parseText(
      await handlers.set_formula({
        filePath,
        sheetName: "费用表",
        cell: "B10",
        formula: "SUM(B2:B9)",
      })
    );
    assert.strictEqual(formulaRes.success, true);

    // Step 6: 读回全表验证
    const after = parseText(
      await handlers.read_range({
        filePath,
        sheetName: "费用表",
      })
    );
    assert.strictEqual(after.success, true);
    // B5 = 3500
    assert.strictEqual(after.values[4][1], 3500, "修改后 B5 应为 3500");
    // A10 = "合计"
    assert.strictEqual(after.values[9][0], "合计");
    // B10 在 reader 中显示为 null（公式无缓存结果）
    assert.strictEqual(after.values[9][1], null,
      "公式格 reader 应返回 null（无缓存结果，设计声明行为）");

    // 用 ExcelJS 验证 B10 公式确实写入
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("费用表")!;
    const b10 = ws.getCell("B10");
    assert.strictEqual(
      typeof b10.value === "object" && b10.value !== null && "formula" in b10.value
        ? (b10.value as any).formula
        : null,
      "SUM(B2:B9)"
    );

    // 其余单元格不变
    assert.strictEqual(after.values[0][0], "项目");
    assert.strictEqual(after.values[0][1], "金额");
    assert.strictEqual(after.values[1][1], 1500); // B2 办公用品
    assert.strictEqual(after.values[2][1], 2500); // B3 差旅费

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S3-02 set_formula 带等号前缀", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "费用表2.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["费用表"] });
    // 写简单数据
    await handlers.write_range({
      filePath,
      sheetName: "费用表",
      startCell: "A1",
      values: [
        ["数值"],
        [10],
        [20],
        [30],
        [40],
      ],
    });

    // 带 `=` 前缀写公式
    const res = parseText(
      await handlers.set_formula({
        filePath,
        sheetName: "费用表",
        cell: "B1",
        formula: "=AVERAGE(A2:A5)",
      })
    );
    assert.strictEqual(res.success, true);

    // ExcelJS 验证去 `=` 后存入
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("费用表")!;
    const cellValue = ws.getCell("B1").value;
    assert.ok(cellValue && typeof cellValue === "object" && "formula" in cellValue);
    assert.strictEqual((cellValue as any).formula, "AVERAGE(A2:A5)");

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S3-03 只改 B5 不影响其他单元格", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "费用表3.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["费用表"] });
    await handlers.write_range({
      filePath,
      sheetName: "费用表",
      startCell: "A1",
      values: [
        ["项目", "金额"],
        ["办公用品", 1500],
        ["差旅费", 2500],
        ["餐饮费", 800],
        ["其他", 3000],
      ],
    });

    // 只改 B5
    await handlers.write_range({
      filePath,
      sheetName: "费用表",
      startCell: "B5",
      values: [[3500]],
    });

    const wb = readWorkbook(filePath);
    // 验证其他行不改
    assert.strictEqual(wb.sheets[0].values[0][0], "项目");
    assert.strictEqual(wb.sheets[0].values[0][1], "金额");
    assert.strictEqual(wb.sheets[0].values[1][1], 1500);
    assert.strictEqual(wb.sheets[0].values[2][1], 2500);
    assert.strictEqual(wb.sheets[0].values[3][1], 800);
    assert.strictEqual(wb.sheets[0].values[4][1], 3500); // B5 已改

    await fs.promises.rm(dir, { recursive: true });
  });
});
