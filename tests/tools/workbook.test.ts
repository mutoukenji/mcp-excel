import { test, describe } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { registerWorkbookTools } from "../../src/tools/workbook.js";

const handlers: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  registerTool: (name: string, _meta: unknown, handler: any) => {
    handlers[name] = handler;
  },
};
registerWorkbookTools(mockServer as any);

function parseText(res: any) {
  return JSON.parse(res.content[0].text);
}

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-wb-"));
}

describe("src/tools/workbook.ts", () => {
  test("WB-01 create_workbook 指定表名", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "wb.xlsx");
    const res = parseText(
      await handlers.create_workbook({ filePath, sheets: ["一月", "二月", "三月"] })
    );
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.sheetsCreated, 3);
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.deepStrictEqual(
      list.sheets.map((s: any) => s.name),
      ["一月", "二月", "三月"]
    );
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-02 不传 sheets", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "wb2.xlsx");
    const res = parseText(await handlers.create_workbook({ filePath }));
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.sheetsCreated, 1);
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.strictEqual(list.sheets[0].name, "Sheet1");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-03 文件已存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "exists.xlsx");
    fs.writeFileSync(filePath, "");
    const result = await handlers.create_workbook({ filePath, sheets: ["A"] });
    const res = parseText(result);
    assert.strictEqual(res.success, false);
    assert.strictEqual(result.isError, true);
    assert.ok(res.error.includes("文件已存在"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-04 扩展名非 .xlsx", async () => {
    const res = parseText(await handlers.create_workbook({ filePath: "a.xls", sheets: ["A"] }));
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("新建文件的扩展名必须是 .xlsx"));
  });

  test("WB-05 sheets 数组内重名", async () => {
    const res = parseText(
      await handlers.create_workbook({ filePath: "a.xlsx", sheets: ["一月", "一月"] })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("重复"));
  });

  test("WB-06 非法表名", async () => {
    const res = parseText(
      await handlers.create_workbook({ filePath: "a.xlsx", sheets: ["a:b"] })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("不合法"));
  });

  test("WB-07 目录不存在", async () => {
    const res = parseText(
      await handlers.create_workbook({
        filePath: path.join(os.tmpdir(), "not-exist-dir", "a.xlsx"),
        sheets: ["A"],
      })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("目录不存在"));
  });

  test("WB-08 list_sheets 正常", async () => {
    const res = parseText(await handlers.list_sheets({ filePath: "./tests/fixtures/sales.xlsx" }));
    assert.strictEqual(res.success, true);
    assert.strictEqual(res.sheets.length, 2);
    assert.strictEqual(res.sheets[0].name, "一月");
  });

  test("WB-09 list_sheets 各格式", async () => {
    const xls = parseText(await handlers.list_sheets({ filePath: "./tests/fixtures/legacy.xls" }));
    assert.strictEqual(xls.success, true);
    const csv = parseText(await handlers.list_sheets({ filePath: "./tests/fixtures/data.csv" }));
    assert.strictEqual(csv.success, true);
  });

  test("WB-10 list_sheets 文件不存在/格式不支持", async () => {
    const notFound = parseText(
      await handlers.list_sheets({ filePath: "./tests/fixtures/not-exist.xlsx" })
    );
    assert.strictEqual(notFound.success, false);
    assert.ok(notFound.error.includes("文件不存在"));
    const unsupported = parseText(
      await handlers.list_sheets({ filePath: "./tests/fixtures/notes.txt" })
    );
    assert.strictEqual(unsupported.success, false);
    assert.ok(unsupported.error.includes("不支持"));
  });

  test("WB-11 add_sheet 正常", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "add.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月", "二月"] });
    const res = parseText(await handlers.add_sheet({ filePath, sheetName: "三月" }));
    assert.strictEqual(res.success, true);
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.ok(list.sheets.some((s: any) => s.name === "三月"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-12 add_sheet 重名", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "add2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月"] });
    const res = parseText(await handlers.add_sheet({ filePath, sheetName: "一月" }));
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "一月" 已存在'));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-13 add_sheet 对 xls / csv", async () => {
    const xls = parseText(
      await handlers.add_sheet({ filePath: "./tests/fixtures/legacy.xls", sheetName: "X" })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
    const csv = parseText(
      await handlers.add_sheet({ filePath: "./tests/fixtures/data.csv", sheetName: "X" })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV"));
  });

  test("WB-14 delete_sheet 正常", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "del.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月", "二月"] });
    const res = parseText(await handlers.delete_sheet({ filePath, sheetName: "二月" }));
    assert.strictEqual(res.success, true);
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.ok(!list.sheets.some((s: any) => s.name === "二月"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-15 delete_sheet 表不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "del2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月"] });
    const res = parseText(await handlers.delete_sheet({ filePath, sheetName: "草稿" }));
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("工作表 \"草稿\" 不存在"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-16 delete_sheet 最后一张", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "del3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月"] });
    const res = parseText(await handlers.delete_sheet({ filePath, sheetName: "一月" }));
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("不能删除最后一张工作表"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-17 delete_sheet 对 xls / csv", async () => {
    const xls = parseText(
      await handlers.delete_sheet({ filePath: "./tests/fixtures/legacy.xls", sheetName: "X" })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
    const csv = parseText(
      await handlers.delete_sheet({ filePath: "./tests/fixtures/data.csv", sheetName: "X" })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV"));
  });

  test("WB-18 rename_sheet 正常", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "ren.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["Sheet1"] });
    const res = parseText(
      await handlers.rename_sheet({ filePath, oldName: "Sheet1", newName: "一季度" })
    );
    assert.strictEqual(res.success, true);
    const list = parseText(await handlers.list_sheets({ filePath }));
    assert.ok(list.sheets.some((s: any) => s.name === "一季度"));
    assert.ok(!list.sheets.some((s: any) => s.name === "Sheet1"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-19 rename_sheet 旧名不存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "ren2.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["一月"] });
    const res = parseText(
      await handlers.rename_sheet({ filePath, oldName: "Sheet1", newName: "X" })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes("工作表 \"Sheet1\" 不存在"));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-20 rename_sheet 新名被占用", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "ren3.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["A", "B"] });
    const res = parseText(
      await handlers.rename_sheet({ filePath, oldName: "A", newName: "B" })
    );
    assert.strictEqual(res.success, false);
    assert.ok(res.error.includes('工作表 "B" 已存在'));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-21 rename_sheet 改成同名", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "ren4.xlsx");
    await handlers.create_workbook({ filePath, sheets: ["A"] });
    const res = parseText(
      await handlers.rename_sheet({ filePath, oldName: "A", newName: "A" })
    );
    assert.strictEqual(res.success, true);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("WB-22 rename_sheet 对 xls / csv", async () => {
    const xls = parseText(
      await handlers.rename_sheet({
        filePath: "./tests/fixtures/legacy.xls",
        oldName: "X",
        newName: "Y",
      })
    );
    assert.strictEqual(xls.success, false);
    assert.ok(xls.error.includes("老版 .xls"));
    const csv = parseText(
      await handlers.rename_sheet({
        filePath: "./tests/fixtures/data.csv",
        oldName: "X",
        newName: "Y",
      })
    );
    assert.strictEqual(csv.success, false);
    assert.ok(csv.error.includes("CSV"));
  });
});
