import { test, describe } from "node:test";
import assert from "node:assert";
import {
  parseCell,
  parseRange,
  toCellName,
  colToLetter,
  letterToCol,
  resolveColumn,
} from "../../src/engine/address.js";
import { ToolError } from "../../src/engine/errors.js";

describe("src/engine/address.ts", () => {
  test("A-01 parseCell B5", () => {
    assert.deepStrictEqual(parseCell("B5"), { row: 5, col: 2 });
  });

  test("A-02 parseCell 大小写不敏感", () => {
    assert.deepStrictEqual(parseCell("b5"), { row: 5, col: 2 });
    assert.deepStrictEqual(parseCell("AB10"), { row: 10, col: 28 });
  });

  test("A-03 parseCell 非法单元格地址", () => {
    for (const input of ["5B", "", "A"]) {
      assert.throws(
        () => parseCell(input),
        (e) => e instanceof ToolError && e.code === "INVALID_CELL"
      );
    }
  });

  test("A-04 parseCell A0 行号为 0", () => {
    assert.throws(
      () => parseCell("A0"),
      (e) => e instanceof ToolError && e.code === "INVALID_CELL"
    );
  });

  test("A-05 parseRange A1:D20", () => {
    const r = parseRange("A1:D20");
    assert.deepStrictEqual(r.start, { row: 1, col: 1 });
    assert.deepStrictEqual(r.end, { row: 20, col: 4 });
  });

  test("A-06 parseRange 颠倒区域自动交换", () => {
    const r = parseRange("D20:A1");
    assert.deepStrictEqual(r.start, { row: 1, col: 1 });
    assert.deepStrictEqual(r.end, { row: 20, col: 4 });
  });

  test("A-07 parseRange 非法区域", () => {
    for (const input of ["A1", "A1:", "A:B"]) {
      assert.throws(
        () => parseRange(input),
        (e) => e instanceof ToolError && e.code === "INVALID_RANGE"
      );
    }
  });

  test("A-08 toCellName / colToLetter", () => {
    assert.strictEqual(toCellName({ row: 5, col: 2 }), "B5");
    assert.strictEqual(colToLetter(1), "A");
    assert.strictEqual(colToLetter(26), "Z");
    assert.strictEqual(colToLetter(27), "AA");
    assert.strictEqual(colToLetter(28), "AB");
  });

  test("A-09 letterToCol", () => {
    assert.strictEqual(letterToCol("a"), 1);
    assert.strictEqual(letterToCol("ab"), 28);
  });

  test("A-10 字母↔列号往返一致", () => {
    for (let col = 1; col <= 702; col++) {
      assert.strictEqual(letterToCol(colToLetter(col)), col);
    }
  });

  test("A-11 resolveColumn 列字母在区域内", () => {
    const range = parseRange("A1:D20");
    assert.strictEqual(resolveColumn("C", range), 3);
  });

  test("A-12 resolveColumn 列字母超出区域", () => {
    const range = parseRange("A1:D20");
    assert.throws(
      () => resolveColumn("E", range),
      (e) =>
        e instanceof ToolError &&
        e.code === "INVALID_PARAMS" &&
        e.message.includes("列 E 不在区域 A1:D20 内")
    );
  });

  test("A-13 resolveColumn 表头名精确匹配", () => {
    const range = parseRange("A1:D20");
    const headerRow = ["姓名", "销售额", "日期"];
    assert.strictEqual(resolveColumn("销售额", range, headerRow), 2);
  });

  test("A-14 resolveColumn 表头名但未传 headerRow", () => {
    const range = parseRange("A1:D20");
    assert.throws(
      () => resolveColumn("销售额", range),
      (e) =>
        e instanceof ToolError &&
        e.code === "INVALID_PARAMS" &&
        e.message.includes("hasHeader: true")
    );
  });

  test("A-15 resolveColumn 表头名不存在", () => {
    const range = parseRange("A1:D20");
    const headerRow = ["姓名", "销售额", "日期"];
    assert.throws(
      () => resolveColumn("销量", range, headerRow),
      (e) =>
        e instanceof ToolError &&
        e.code === "INVALID_PARAMS" &&
        e.message.includes("姓名") &&
        e.message.includes("销售额") &&
        e.message.includes("日期")
    );
  });

  test("A-16 区域不从 A 列开始时表头偏移", () => {
    const range = parseRange("C1:F10");
    const headerRow = ["姓名", "销售额"];
    assert.strictEqual(resolveColumn("销售额", range, headerRow), 4);
  });
});
