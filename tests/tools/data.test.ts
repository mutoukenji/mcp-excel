import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { registerDataTools } from "../../src/tools/data.js";
import { registerWriteTools } from "../../src/tools/write.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";
import { registerReadTools } from "../../src/tools/read.js";
import { readWorkbook } from "../../src/engine/reader.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerDataTools(mockServer as any);
registerWriteTools(mockServer as any);
registerWorkbookTools(mockServer as any);
registerReadTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-data-"));
}

describe("src/tools/data.ts", () => {
  test("D-01 按列字母升序排数字", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["a", 3],
        ["b", 1],
        ["c", 2],
      ],
    });
    const res = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B3",
        keyColumn: "B",
        order: "asc",
      })
    );
    assert.strictEqual(res.success, true);
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values, [
      ["b", 1],
      ["c", 2],
      ["a", 3],
    ]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-02 按表头名降序", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额"],
        ["张三", 100],
        ["李四", 200],
      ],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:B3",
      keyColumn: "销售额",
      order: "desc",
      hasHeader: true,
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "姓名");
    assert.strictEqual(wb.sheets[0].values[1][0], "李四");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-03 null 永远排最后", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[null], [1], [2], ["keep"]],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A3",
      keyColumn: "A",
      order: "asc",
    });
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values.map((r: any) => r[0]), [1, 2, null, "keep"]);
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A3",
      keyColumn: "A",
      order: "desc",
    });
    const wb2 = readWorkbook(filePath);
    assert.deepStrictEqual(wb2.sheets[0].values.map((r: any) => r[0]), [2, 1, null, "keep"]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-04 稳定排序", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        [1, "a"],
        [1, "b"],
        [1, "c"],
      ],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:B3",
      keyColumn: "A",
      order: "asc",
    });
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values, [
      [1, "a"],
      [1, "b"],
      [1, "c"],
    ]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-05 字符串按中文排序", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort5.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["王"], ["张"], ["李"]],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A3",
      keyColumn: "A",
      order: "asc",
    });
    const wb = readWorkbook(filePath);
    const names = wb.sheets[0].values.map((r: any) => r[0]);
    assert.deepStrictEqual(names, ["李", "王", "张"]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-06 数字与字符串混排", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort6.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["x"], [1], ["a"], [2]],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A4",
      keyColumn: "A",
      order: "asc",
    });
    const wb = readWorkbook(filePath);
    // 数字之间按数值，其他按字符串序
    const values = wb.sheets[0].values.map((r: any) => r[0]);
    assert.strictEqual(values[0], 1);
    assert.strictEqual(values[1], 2);
    assert.strictEqual(typeof values[2], "string");
    assert.strictEqual(typeof values[3], "string");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-07 列字母超出区域", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort7.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[1, 2, 3]],
    });
    const res = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B1",
        keyColumn: "E",
        order: "asc",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("列 E 不在区域"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-08 表头名打错", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort8.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["姓名", "销售额"]],
    });
    const res = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B1",
        keyColumn: "销量",
        order: "asc",
        hasHeader: true,
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("销量"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-09 用表头名但 hasHeader 未开", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort9.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["姓名", "销售额"]],
    });
    const res = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B1",
        keyColumn: "销售额",
        order: "asc",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("hasHeader: true"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-10 区域外内容不动", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sort10.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        [3, "x"],
        [1, "x"],
        [2, "x"],
        ["out", "side"],
      ],
    });
    await handlers.sort_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:B3",
      keyColumn: "A",
      order: "asc",
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[3][0], "out");
    assert.strictEqual(wb.sheets[0].values[3][1], "side");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-11 csv 排序可用", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "sortcsv.csv");
    fs.copyFileSync("./tests/fixtures/data.csv", filePath);
    const res = parseText(
      await handlers.sort_range({
        filePath,
        sheetName: "任意",
        range: "A1:C3",
        keyColumn: "B",
        order: "asc",
      })
    );
    assert.strictEqual(res.success, true);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-12 对 xls 排序", async () => {
    const res = parseText(
      await handlers.sort_range({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "Sheet1",
        range: "A1:D6",
        keyColumn: "A",
        order: "asc",
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("老版 .xls"));
  });

  test("D-13 filter_range 数值比较", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额"],
        ["张三", 800],
        ["李四", 1200],
      ],
    });
    const res = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B3",
        conditions: [{ column: "销售额", op: ">", value: 1000 }],
        hasHeader: true,
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 2);
    assert.strictEqual(res.values[1][0], "李四");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-14 多条件并且", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["a", 1, 10],
        ["b", 2, 5],
        ["c", 1, 15],
      ],
    });
    const res = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:C3",
        conditions: [
          { column: "B", op: "=", value: 1 },
          { column: "C", op: ">", value: 10 },
        ],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 1);
    assert.strictEqual(res.values[0][0], "c");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-15 宽松相等", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[1000], ["1000"]],
    });
    const res = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A2",
        conditions: [{ column: "A", op: "=", value: "1000" }],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 2);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-16 != / >= / <= / <", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[1], [2], [3]],
    });
    for (const [op, value, expected] of [
      ["!=", 2, 2],
      [">=", 2, 2],
      ["<=", 2, 2],
      ["<", 2, 1],
    ] as const) {
      const res = parseText(
        await handlers.filter_range({
          filePath,
          sheetName: "Sheet1",
          range: "A1:A3",
          conditions: [{ column: "A", op, value }],
        })
      );
      assert.strictEqual(res.values.length, expected);
    }
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-17 contains 大小写敏感", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter5.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["Apple"], ["app"]],
    });
    const res = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A2",
        conditions: [{ column: "A", op: "contains", value: "app" }],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 1);
    assert.strictEqual(res.values[0][0], "app");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-18 字符串间的大小比较", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter6.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["b"], ["a"]],
    });
    const res = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A2",
        conditions: [{ column: "A", op: ">", value: "a" }],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.values.length, 1);
    assert.strictEqual(res.values[0][0], "b");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-19 0 行匹配", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter7.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["姓名", "销售额"],
        ["张三", 800],
      ],
    });
    const withHeader = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:B2",
        conditions: [{ column: "销售额", op: ">", value: 1000 }],
        hasHeader: true,
      })
    );
    assert.deepStrictEqual(withHeader.values, [["姓名", "销售额"]]);
    const noHeader = parseText(
      await handlers.filter_range({
        filePath,
        sheetName: "Sheet1",
        range: "A2:B2",
        conditions: [{ column: "B", op: ">", value: 1000 }],
      })
    );
    assert.deepStrictEqual(noHeader.values, []);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-20 不改文件", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "filter8.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[1], [2]],
    });
    const before = fs.readFileSync(filePath);
    await handlers.filter_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A2",
      conditions: [{ column: "A", op: ">", value: 0 }],
    });
    const after = fs.readFileSync(filePath);
    assert.deepStrictEqual(before, after);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-21 xls 允许筛选", async () => {
    const res = parseText(
      await handlers.filter_range({
        filePath: "./tests/fixtures/legacy.xls",
        sheetName: "一月",
        range: "A1:D6",
        conditions: [{ column: "A", op: "=", value: "张三" }],
      })
    );
    assert.strictEqual(res.success, true);
  });

  test("D-22 dedupe_range 基本去重", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["a"], ["b"], ["a"], ["c"]],
    });
    const res = parseText(
      await handlers.dedupe_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A4",
        keyColumns: ["A"],
      })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.removedCount, 1);
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values.map((r: any) => r[0]), ["a", "b", "c"]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-23 多键组合", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["a", 1],
        ["a", 2],
        ["a", 1],
      ],
    });
    await handlers.dedupe_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:B3",
      keyColumns: ["A", "B"],
    });
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values, [
      ["a", 1],
      ["a", 2],
    ]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-24 尾部置空不移动区域外", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [
        ["a"],
        ["a"],
        ["b"],
        ["out"],
      ],
    });
    await handlers.dedupe_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A3",
      keyColumns: ["A"],
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[3][0], "out");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-25 hasHeader 参与", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["姓名"], ["a"], ["a"]],
    });
    await handlers.dedupe_range({
      filePath,
      sheetName: "Sheet1",
      range: "A1:A3",
      keyColumns: ["A"],
      hasHeader: true,
    });
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets[0].values[0][0], "姓名");
    assert.strictEqual(wb.sheets[0].values[1][0], "a");
    assert.strictEqual(wb.sheets[0].values.length, 2);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-26 无重复", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe5.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [["a"], ["b"], ["c"]],
    });
    const res = parseText(
      await handlers.dedupe_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A3",
        keyColumns: ["A"],
      })
    );
    assert.strictEqual(res.removedCount, 0);
    const wb = readWorkbook(filePath);
    assert.deepStrictEqual(wb.sheets[0].values.map((r: any) => r[0]), ["a", "b", "c"]);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("D-27 归一化后判重", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "dedupe6.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    await handlers.write_range({
      filePath,
      sheetName: "Sheet1",
      startCell: "A1",
      values: [[1], ["1"]],
    });
    const res = parseText(
      await handlers.dedupe_range({
        filePath,
        sheetName: "Sheet1",
        range: "A1:A2",
        keyColumns: ["A"],
      })
    );
    // 类型不同不算重复，removedCount 应为 0
    assert.strictEqual(res.removedCount, 0);
    await fs.promises.rm(dir, { recursive: true });
  });
});
