import { test, describe, mock } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import ExcelJS from "exceljs";
import {
  createWorkbook,
  editWorkbook,
  normalizeCellValue,
} from "../../src/engine/writer.js";
import { readWorkbook } from "../../src/engine/reader.js";
import { ToolError } from "../../src/engine/errors.js";

async function tmpDir(): Promise<string> {
  return await fs.promises.mkdtemp(path.join(os.tmpdir(), "mcp-excel-writer-"));
}

describe("src/engine/writer.ts", () => {
  test("W-01 createWorkbook 新建多表", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "new.xlsx");
    await createWorkbook(filePath, ["A", "B", "C"]);
    assert.ok(fs.existsSync(filePath));
    const wb = readWorkbook(filePath);
    assert.strictEqual(wb.sheets.length, 3);
    assert.deepStrictEqual(
      wb.sheets.map((s) => s.name),
      ["A", "B", "C"]
    );
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-02 目标已存在", async () => {
    const dir = await tmpDir();
    const filePath = path.join(dir, "exists.xlsx");
    await createWorkbook(filePath, ["Sheet1"]);
    const before = fs.readFileSync(filePath);
    await assert.rejects(
      async () => createWorkbook(filePath, ["X"]),
      (e) => e instanceof ToolError && e.code === "FILE_EXISTS"
    );
    const after = fs.readFileSync(filePath);
    assert.deepStrictEqual(before, after);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-03 目录不存在", async () => {
    const filePath = path.join(os.tmpdir(), "not-exist-dir", "new.xlsx");
    await assert.rejects(
      async () => createWorkbook(filePath, ["Sheet1"]),
      (e) => e instanceof ToolError && e.code === "DIR_NOT_FOUND"
    );
  });

  test("W-04 editWorkbook 修改并保存", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "src.xlsx");
    await createWorkbook(src, ["Sheet1"]);
    await editWorkbook(src, (wb) => {
      wb.getWorksheet("Sheet1")!.getCell("B5").value = 3500;
    });
    const wb = readWorkbook(src);
    assert.strictEqual(wb.sheets[0].values[4][1], 3500);
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-05 编辑不存在文件", async () => {
    await assert.rejects(
      async () => editWorkbook(path.join(os.tmpdir(), "not-exist.xlsx"), () => {}),
      (e) => e instanceof ToolError && e.code === "FILE_NOT_FOUND"
    );
  });

  test("W-06 编辑 .xls", async () => {
    await assert.rejects(
      async () => editWorkbook("./tests/fixtures/legacy.xls", () => {}),
      (e) => e instanceof ToolError && e.code === "READ_ONLY_FORMAT"
    );
  });

  test("W-07 编辑 csv", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "data.csv");
    fs.copyFileSync("./tests/fixtures/data.csv", src);
    await editWorkbook(src, (wb) => {
      wb.worksheets[0].getCell("A2").value = "王五";
    });
    const wb = readWorkbook(src);
    assert.strictEqual(wb.sheets[0].values[1][0], "王五");
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-08 不支持格式", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "notes.txt");
    fs.writeFileSync(src, "not a spreadsheet");
    await assert.rejects(
      async () => editWorkbook(src, () => {}),
      (e) => e instanceof ToolError && e.code === "UNSUPPORTED_FORMAT"
    );
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-09 损坏文件", async () => {
    await assert.rejects(
      async () => editWorkbook("./tests/fixtures/corrupt.xlsx", () => {}),
      (e) => e instanceof ToolError && e.code === "UNKNOWN"
    );
  });

  test("W-10 原子写 mutator 中途抛错", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "atomic.xlsx");
    await createWorkbook(src, ["Sheet1"]);
    fs.writeFileSync(src + ".orig", fs.readFileSync(src));
    const before = fs.readFileSync(src);
    await assert.rejects(
      async () =>
        editWorkbook(src, (wb) => {
          wb.getWorksheet("Sheet1")!.getCell("A1").value = 1;
          throw new ToolError("UNKNOWN", "mutator boom");
        }),
      (e) => e instanceof ToolError && e.message === "mutator boom"
    );
    const after = fs.readFileSync(src);
    assert.deepStrictEqual(before, after);
    assert.ok(!fs.existsSync(src + ".tmp-" + process.pid));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-11 原子写无残留", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "clean.xlsx");
    await createWorkbook(src, ["Sheet1"]);
    await editWorkbook(src, (wb) => {
      wb.getWorksheet("Sheet1")!.getCell("A1").value = 1;
    });
    assert.ok(!fs.existsSync(src + ".tmp-" + process.pid));
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-12 文件被占用", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "busy.xlsx");
    await createWorkbook(src, ["Sheet1"]);
    const fd = fs.openSync(src, "r+");
    try {
      await assert.rejects(
        async () =>
          editWorkbook(src, (wb) => {
            wb.getWorksheet("Sheet1")!.getCell("A1").value = 1;
          }),
        (e) => e instanceof ToolError && e.code === "FILE_BUSY"
      );
    } finally {
      fs.closeSync(fd);
    }
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-13 mutator 抛非 ToolError", async () => {
    const dir = await tmpDir();
    const src = path.join(dir, "plain-error.xlsx");
    await createWorkbook(src, ["Sheet1"]);
    await assert.rejects(
      async () =>
        editWorkbook(src, () => {
          throw new Error("英文库错误");
        }),
      (e) => e instanceof ToolError && e.code === "UNKNOWN"
    );
    await fs.promises.rm(dir, { recursive: true });
  });

  test("W-14 normalizeCellValue 全映射", () => {
    assert.strictEqual(normalizeCellValue(null), null);
    assert.strictEqual(normalizeCellValue(undefined), null);
    assert.strictEqual(normalizeCellValue("a"), "a");
    assert.strictEqual(normalizeCellValue(1), 1);
    assert.strictEqual(normalizeCellValue(true), true);
    const d = new Date("2026-01-05T00:00:00Z");
    assert.strictEqual(normalizeCellValue(d), d.toISOString());
    assert.strictEqual(
      normalizeCellValue({ richText: [{ text: "a" }, { text: "b" }] } as any),
      "ab"
    );
    assert.strictEqual(
      normalizeCellValue({ text: "link", hyperlink: "http://x" } as any),
      "link"
    );
    assert.strictEqual(
      normalizeCellValue({ formula: "A1+1", result: 2 } as any),
      2
    );
    assert.strictEqual(
      normalizeCellValue({ formula: "A1+1" } as any),
      null
    );
    assert.strictEqual(normalizeCellValue({ error: "#VALUE!" } as any), null);
    assert.strictEqual(normalizeCellValue({ unknown: 1 } as any), null);
  });
});
