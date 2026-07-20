/**
 * S5 数据加工流水线（write → dedupe → sort → format → merge → filter）
 *
 * 涉及模块：6 个工具文件 + 全部引擎文件（覆盖工具间对同一文件的接力修改）
 * 步骤：
 *   1. create_workbook + write_range 写入 20 行含重复、未排序、无样式的数据；
 *   2. dedupe_range(keyColumns=["A"], hasHeader=true)；
 *   3. sort_range(keyColumn="销售额", order="desc", hasHeader=true)；
 *   4. format_cells 美化标题行；merge_cells("A1:D1") 合并标题；
 *   5. filter_range(销售额 > 1000) 取结果；
 *   6. read_range 读全表终态验证。
 * 预期：每一步都基于上一步的落盘结果正确执行；终态文件 = 已去重 + 已排序 + 有样式 + 有合并；
 *   filter 结果与终态文件内容一致。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import ExcelJS from "exceljs";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerReadTools } from "../../src/tools/read.js";
import { registerDataTools } from "../../src/tools/data.js";
import { registerFormatTools } from "../../src/tools/format.js";
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
registerDataTools(mockServer as any);
registerFormatTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-s5-"));
}

describe("S5 数据加工流水线", () => {
  test("S5-01 写入含重复数据的初始表", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });

    // 20 行数据，含重复（张三 x2, 王五 x2, 赵六 x2），未排序
    const writeRes = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [
          ["姓名", "销售额", "日期", "备注"],
          ["张三", 1000, "2026-01-01", null],
          ["李四", 800, "2026-01-02", "新客户"],
          ["王五", 1500, "2026-01-03", null],
          ["赵六", 600, "2026-01-04", "已回款"],
          ["张三", 1000, "2026-01-01", null],  // 重复：张三
          ["钱七", 2000, "2026-01-05", null],
          ["孙八", 300, "2026-01-06", "待确认"],
          ["王五", 1500, "2026-01-03", null],  // 重复：王五
          ["周九", 1200, "2026-01-07", null],
          ["吴十", 450, "2026-01-08", "已回款"],
          ["赵六", 600, "2026-01-04", "已回款"], // 重复：赵六
          ["郑一", 1800, "2026-01-09", null],
          ["陈二", 900, "2026-01-10", null],
          ["刘三", 700, "2026-01-11", "新客户"],
          ["黄四", 1100, "2026-01-12", null],
          ["林五", 350, "2026-01-13", null],
          ["何六", 1600, "2026-01-14", "待确认"],
          ["马七", 550, "2026-01-15", null],
          ["罗八", 1300, "2026-01-16", "已回款"],
          ["高九", 250, "2026-01-17", null],
        ],
      })
    );
    assert.strictEqual(writeRes.success, true);
    assert.strictEqual(writeRes.cellsWritten, 84); // 21 rows * 4 cols

    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].rowCount, 21);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S5-02 去重：按姓名列去重保留首现", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额", "日期", "备注"],
        ["张三", 1000, "2026-01-01", null],
        ["李四", 800, "2026-01-02", "新客户"],
        ["王五", 1500, "2026-01-03", null],
        ["赵六", 600, "2026-01-04", "已回款"],
        ["张三", 1000, "2026-01-01", null],  // 重复
        ["钱七", 2000, "2026-01-05", null],
        ["孙八", 300, "2026-01-06", "待确认"],
        ["王五", 1500, "2026-01-03", null],  // 重复
        ["周九", 1200, "2026-01-07", null],
        ["吴十", 450, "2026-01-08", "已回款"],
        ["赵六", 600, "2026-01-04", "已回款"], // 重复
        ["郑一", 1800, "2026-01-09", null],
        ["陈二", 900, "2026-01-10", null],
        ["刘三", 700, "2026-01-11", "新客户"],
        ["黄四", 1100, "2026-01-12", null],
        ["林五", 350, "2026-01-13", null],
        ["何六", 1600, "2026-01-14", "待确认"],
        ["马七", 550, "2026-01-15", null],
        ["罗八", 1300, "2026-01-16", "已回款"],
        ["高九", 250, "2026-01-17", null],
      ],
    });

    // Step 2: 去重（按姓名）
    const dedupeRes = parseText(
      await handlers.dedupe_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:D21",
        keyColumns: ["A"],
        hasHeader: true,
      })
    );
    assert.strictEqual(dedupeRes.success, true);
    // 21 rows - 1 header = 20 data rows, 3 duplicates removed = 17 data + 1 header = 18 non-null rows
    assert.ok(dedupeRes.removedCount >= 3, `去重数应为 >= 3，实际: ${dedupeRes.removedCount}`);

    // 验证表头未被删除
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "姓名");

    // 统计非空行数量
    const nonNullRows = wb.sheets[0].values.filter(
      (r: any) => !r.every((v: any) => v === null)
    );
    // 去重后应有 1 header + 17 unique data rows = 18 non-null rows
    assert.strictEqual(nonNullRows.length, 18);

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S5-03 排序：按销售额降序（表头不动）", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额", "日期", "备注"],
        ["张三", 1000, "2026-01-01", null],
        ["李四", 800, "2026-01-02", "新客户"],
        ["王五", 1500, "2026-01-03", null],
        ["赵六", 600, "2026-01-04", "已回款"],
        ["钱七", 2000, "2026-01-05", null],
      ],
    });

    // 先 dedupe 再 sort（range 覆盖全部 6 行）
    await handlers.dedupe_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:D6",
      keyColumns: ["A"],
      hasHeader: true,
    });

    const sortRes = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:D6",
        keyColumn: "B",  // 销售额列
        order: "desc",
        hasHeader: true,
      })
    );
    assert.strictEqual(sortRes.success, true);

    const wb = readWorkbook(filePath);
    const vals = wb.sheets[0].values;
    // 表头应在第一行
    assert.strictEqual(vals[0][0], "姓名");
    // 按销售额降序：钱七(2000), 王五(1500), 张三(1000), 李四(800), 赵六(600)
    assert.strictEqual(vals[1][0], "钱七");
    assert.strictEqual(vals[1][1], 2000);
    assert.strictEqual(vals[2][0], "王五");
    assert.strictEqual(vals[3][0], "张三");
    assert.strictEqual(vals[4][0], "李四");

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S5-04 美化标题行 + 合并单元格", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额", "日期", "备注"],
        ["张三", 1000, "2026-01-01", null],
        ["李四", 800, "2026-01-02", "新客户"],
      ],
    });

    // format_cells: 美化标题行 A1:D1
    const fmtRes = parseText(
      await handlers.format_cells({
        filePath,
        sheetName: "Sheet1",
        range: "A1:D1",
        style: {
          font: { bold: true, color: "FFFFFF", size: 12 },
          fill: { color: "305496" },
          alignment: { horizontal: "center", vertical: "middle" },
          border: { style: "thin", color: "999999" },
        },
      })
    );
    assert.strictEqual(fmtRes.success, true);

    // merge_cells: 合并 A1:D1
    const mergeRes = parseText(
      await handlers.merge_cells({
        filePath,
        sheetName: "Sheet1",
        range: "A1:D1",
      })
    );
    assert.strictEqual(mergeRes.success, true);

    // ExcelJS 验证
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet("Sheet1")!;

    // 检查 A1 字体
    const a1 = ws.getCell("A1");
    assert.strictEqual(a1.font.bold, true);
    assert.strictEqual((a1.font.color as any).argb, "FFFFFFFF");
    assert.strictEqual(a1.fill?.type, "pattern");

    // 检查合并区域
    const mergedCells = (ws as any)._merges || (ws as any).model?.merges || [];
    assert.ok(mergedCells.length > 0 || ws.getCell("B1").isMerged,
      "应有合并单元格");

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S5-05 filter_range 筛选销售额 > 1000", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额", "日期", "备注"],
        ["张三", 1200, "2026-01-01", null],
        ["李四", 800, "2026-01-02", "新客户"],
        ["王五", 1500, "2026-01-03", null],
        ["赵六", 600, "2026-01-04", "已回款"],
        ["钱七", 2000, "2026-01-05", null],
      ],
    });

    // 筛选销售额 > 1000
    const filterRes = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:D6",
        hasHeader: true,
        conditions: [{ column: "销售额", op: ">", value: 1000 }],
      })
    );
    assert.strictEqual(filterRes.success, true);
    // 表头 + 销售额 > 1000 的行：张三(1200), 王五(1500), 钱七(2000) = 4 行
    assert.strictEqual(filterRes.values.length, 4,
      "表头 + 3 条匹配数据 = 4 行");
    assert.strictEqual(filterRes.values[0][0], "姓名"); // 表头
    // 找到销售额 > 1000 的人
    const matchedNames = filterRes.values.slice(1).map((r: any) => r[0]);
    assert.ok(matchedNames.includes("张三"));
    assert.ok(matchedNames.includes("王五"));
    assert.ok(matchedNames.includes("钱七"));
    assert.ok(!matchedNames.includes("李四"));  // 800 不应匹配
    assert.ok(!matchedNames.includes("赵六"));  // 600 不应匹配

    await fs.promises.rm(dir, { recursive: true });
  });

  test("S5-06 完整流水线端到端", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "pipeline.xlsx");

    // Step 1: 创建并写数据
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const writeRes = parseText(
      await handlers.write_range({
        filePath,
        sheetName: "Sheet1",
        startCell: "A1",
        values: [
          ["姓名", "销售额", "日期", "备注"],
          ["张三", 1200, "2026-01-01", null],
          ["李四", 800, "2026-01-02", "新客户"],
          ["王五", 1500, "2026-01-03", null],
          ["赵六", 600, "2026-01-04", "已回款"],
          ["张三", 1200, "2026-01-01", null],  // dup
          ["钱七", 2000, "2026-01-05", null],
          ["孙八", 300, "2026-01-06", "待确认"],
          ["王五", 1500, "2026-01-03", null],  // dup
        ],
      })
    );
    assert.strictEqual(writeRes.success, true);

    // Step 2: dedupe（range 需包含全部 9 行）
    const dedupeRes = parseText(
      await handlers.dedupe_range({
        filePath, sheetName: "Sheet1",
        range: "A1:D9", keyColumns: ["A"], hasHeader: true,
      })
    );
    assert.strictEqual(dedupeRes.success, true);
    assert.strictEqual(dedupeRes.removedCount, 2); // 张三 dup + 王五 dup

    // Step 3: sort desc（range 同 A1:D9）
    const sortRes = parseText(
      await handlers.sort_range({
        filePath, sheetName: "Sheet1",
        range: "A1:D9", keyColumn: "B", order: "desc", hasHeader: true,
      })
    );
    assert.strictEqual(sortRes.success, true);

    // Step 4: format header（加粗+填充）
    const fmtRes = parseText(
      await handlers.format_cells({
        filePath, sheetName: "Sheet1", range: "A1:D1",
        style: { font: { bold: true }, fill: { color: "305496" } },
      })
    );
    assert.strictEqual(fmtRes.success, true);

    // Step 5: filter 销售额 > 1000（在 merge 之前执行，保持表头可读）
    const filterRes = parseText(
      await handlers.filter_range({
        filePath, sheetName: "Sheet1",
        range: "A1:D9", hasHeader: true,
        conditions: [{ column: "销售额", op: ">", value: 1000 }],
      })
    );
    assert.strictEqual(filterRes.success, true);
    const matchedNames = filterRes.values.slice(1).map((r: any) => r[0]);
    // 去重+排序后，unique names > 1000: 钱七(2000), 王五(1500), 张三(1200)
    assert.ok(matchedNames.includes("钱七"));
    assert.ok(matchedNames.includes("王五"));
    assert.ok(matchedNames.includes("张三"));
    assert.strictEqual(matchedNames.length, 3);

    // Step 6: merge A1:D1（作为最后一步显示美化，之后不再读表头）
    const mergeRes = parseText(
      await handlers.merge_cells({
        filePath, sheetName: "Sheet1", range: "A1:D1",
      })
    );
    assert.strictEqual(mergeRes.success, true);

    // Step 7: read_range 读全表验证终态
    const finalRead = parseText(
      await handlers.read_range({ filePath, sheetName: "Sheet1" })
    );
    assert.strictEqual(finalRead.success, true);
    // 表头在首行（合并后首格保留值："姓名"）
    assert.strictEqual(finalRead.values[0][0], "姓名");
    // 非空数据行（去重后 6 unique + 1 header = 7, 尾部 2 null rows from dedupe trimmed by reader）
    const nonNull = finalRead.values.filter(
      (r: any) => !r.every((v: any) => v === null)
    );
    assert.strictEqual(nonNull.length, 7, `应有 7 行非空数据（1 header + 6 data）`);

    // ExcelJS 验证格式和合并均生效
    const xlWb = new ExcelJS.Workbook();
    await xlWb.xlsx.readFile(filePath);
    const ws = xlWb.getWorksheet("Sheet1")!;
    assert.strictEqual(ws.getCell("A1").font.bold, true, "表头应加粗");
    // 验证有合并单元格
    const hasMerge = (ws as any)._merges?.length > 0 || ws.getCell("B1").isMerged;
    assert.ok(hasMerge, "应有合并单元格");

    await fs.promises.rm(dir, { recursive: true });
  });
});
