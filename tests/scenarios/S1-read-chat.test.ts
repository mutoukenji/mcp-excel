/**
 * S1 读表问答（需求场景一 / 架构流程一）
 *
 * 涉及模块：tools/read.ts、tools/workbook.ts(list_sheets)、engine/reader.ts、engine/address.ts
 * 步骤：
 *   1. 对 sales.xlsx 调 list_sheets 确认"一月"存在；
 *   2. 调 read_range(filePath, sheetName="一月")；
 *   3. 模拟 AI 行为：在返回的二维数组中定位"销售额"列并求和。
 * 预期：两步均 success: true；返回值结构与设计 3.1 节示例一致（日期为 ISO 字符串、空格为 null）；求和结果与 fixture 预置总额一致。
 */
import { test, describe } from "node:test";
import assert from "node:assert";
import { registerReadTools } from "../../src/tools/read.js";
import { registerWorkbookTools } from "../../src/tools/workbook.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerReadTools(mockServer as any);
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

/** 模拟 AI 定位"销售额"列并求和的完整链路 */
function simulateAISummarize(values: any[][]): {
  totalSales: number;
  salesRowCount: number;
  dateIsISO: boolean;
  nullsCorrectlyHandled: boolean;
} {
  // Step 1: 定位表头行，找到"销售额"列的索引
  const headerRow = values[0];
  const salesColIdx = headerRow.indexOf("销售额");

  // Step 2: 提取所有数据行的销售额（跳过表头行，跳过 null 行）
  let totalSales = 0;
  let salesRowCount = 0;
  let dateIsISO = true;
  let nullsCorrectlyHandled = true;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    // 跳过全 null 的空行
    if (row.every((v: any) => v === null)) continue;

    const salesValue = row[salesColIdx];
    if (salesValue !== null) {
      totalSales += Number(salesValue);
      salesRowCount++;
    }

    // 验证日期列为 ISO 字符串
    const dateVal = row[headerRow.indexOf("日期")];
    if (dateVal !== null && typeof dateVal === "string") {
      if (!dateVal.startsWith("2026-01-")) dateIsISO = false;
    }

    // 验证 null 单元格被正确保留
    const notesVal = row[headerRow.indexOf("备注")];
    // 有些行备注是 null，有些有值，这都 OK
  }

  return { totalSales, salesRowCount, dateIsISO, nullsCorrectlyHandled };
}

describe("S1 读表问答", () => {
  test("S1-01 list_sheets 确认'一月'存在", async () => {
    const res = parseText(
      await handlers.list_sheets({ filePath: "./tests/fixtures/sales.xlsx" })
    );
    assert.strictEqual(res.success, true);
    assert.ok(Array.isArray(res.sheets));
    assert.strictEqual(res.sheets.length, 2);

    const jan = res.sheets.find((s: any) => s.name === "一月");
    assert.ok(jan, "应存在'一月'表");
    assert.strictEqual(jan.rowCount, 6); // 5 data rows + 1 middle blank row, 2 trailing trimmed
    assert.strictEqual(jan.colCount, 4);
  });

  test("S1-02 read_range 读整张'一月'表", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );
    assert.strictEqual(res.success, true);
    assert.ok(Array.isArray(res.values));
    assert.ok(res.values.length > 0);

    // 验证表头
    assert.strictEqual(res.values[0][0], "姓名");
    assert.strictEqual(res.values[0][1], "销售额");
    assert.strictEqual(res.values[0][2], "日期");
    assert.strictEqual(res.values[0][3], "备注");
  });

  test("S1-03 日期为 ISO 字符串、空格为 null", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );

    // 第 2 行（张三）：日期应为 ISO 字符串，备注应为 null
    const row2 = res.values[1];
    assert.strictEqual(row2[0], "张三");
    assert.strictEqual(row2[1], 1200);
    assert.ok(typeof row2[2] === "string" && row2[2].startsWith("2026-01-"));
    assert.strictEqual(row2[3], null);

    // 第 5 行（中间空白行）：全部为 null 且被保留
    const row5 = res.values[4];
    assert.deepStrictEqual(row5, [null, null, null, null]);
  });

  test("S1-04 模拟 AI：定位'销售额'列并求和", async () => {
    const res = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );

    const { totalSales, salesRowCount } = simulateAISummarize(res.values);

    // Fixture 中销售额：张三=1200, 李四=980.5, 王五=1500, 赵六=2000 = 5680.5
    assert.strictEqual(totalSales, 5680.5);
    assert.strictEqual(salesRowCount, 4);
  });

  test("S1-05 read_range 不传 sheetName 默认第一张表", async () => {
    const res = parseText(
      await handlers.read_range({ filePath: "./tests/fixtures/sales.xlsx" })
    );
    assert.strictEqual(res.success, true);
    // 第一张表是"一月"
    assert.strictEqual(res.values[0][0], "姓名");
  });

  test("S1-06 结构一致性：list_sheets 的行列数与 read_range 实际值匹配", async () => {
    const list = parseText(
      await handlers.list_sheets({ filePath: "./tests/fixtures/sales.xlsx" })
    );
    const janList = list.sheets.find((s: any) => s.name === "一月");
    assert.ok(janList);

    const data = parseText(
      await handlers.read_range({
        filePath: "./tests/fixtures/sales.xlsx",
        sheetName: "一月",
      })
    );

    // rowCount 应与 values.length 匹配
    assert.strictEqual(data.values.length, janList.rowCount,
      `list_sheets rowCount=${janList.rowCount} 应与 read_range values.length=${data.values.length} 一致`);
    // colCount：每行长度都等于 colCount
    for (const row of data.values) {
      assert.strictEqual(row.length, janList.colCount);
    }
  });
});
