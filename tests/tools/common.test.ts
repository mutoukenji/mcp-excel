import { test, describe } from "node:test";
import assert from "node:assert";
import {
  ok,
  fail,
  tool,
  filePathSchema,
  sheetNameSchema,
  cellSchema,
  rangeSchema,
  valuesSchema,
} from "../../src/tools/common.js";
import { ToolError } from "../../src/engine/errors.js";

describe("src/tools/common.ts", () => {
  test("C-01 ok 成功包络", () => {
    const result = ok({ a: 1 });
    assert.strictEqual(result.content[0].type, "text");
    const parsed = JSON.parse(result.content[0].text);
    assert.strictEqual(parsed.success, true);
    assert.strictEqual(parsed.a, 1);
    assert.strictEqual(result.isError, undefined);
  });

  test("C-02 fail 失败包络", () => {
    const e = new ToolError("SHEET_EXISTS", '工作表 "四月" 已存在。');
    const result = fail(e);
    const parsed = JSON.parse(result.content[0].text);
    assert.strictEqual(parsed.success, false);
    assert.ok(parsed.error.includes("已存在"));
    assert.strictEqual(result.isError, true);
  });

  test("C-03 tool 成功路径", async () => {
    const handler = tool(async () => ({ x: 2 }));
    const result = await handler({});
    const parsed = JSON.parse(result.content[0].text);
    assert.strictEqual(parsed.success, true);
    assert.strictEqual(parsed.x, 2);
  });

  test("C-04 tool 吞异常", async () => {
    for (const error of [
      new ToolError("UNKNOWN", "x"),
      new Error("boom"),
      "string error",
    ]) {
      const handler = tool(() => {
        throw error;
      });
      const result = await handler({});
      const parsed = JSON.parse(result.content[0].text);
      assert.strictEqual(parsed.success, false);
      assert.strictEqual(result.isError, true);
    }
  });

  test("C-05 filePathSchema", () => {
    assert.strictEqual(filePathSchema.safeParse("").success, false);
    assert.strictEqual(filePathSchema.safeParse("C:\\a.xlsx").success, true);
  });

  test("C-06 sheetNameSchema 边界", () => {
    const name31 = "a".repeat(31);
    assert.strictEqual(sheetNameSchema.safeParse(name31).success, true);
    assert.strictEqual(sheetNameSchema.safeParse("a".repeat(32)).success, false);
    for (const bad of ["a[b", "a]b", "a:b", "a*b", "a?b", "a/b", "a\\b"]) {
      assert.strictEqual(sheetNameSchema.safeParse(bad).success, false);
    }
  });

  test("C-07 cellSchema / rangeSchema", () => {
    assert.strictEqual(cellSchema.safeParse("B5").success, true);
    assert.strictEqual(cellSchema.safeParse("5B").success, false);
    assert.strictEqual(rangeSchema.safeParse("A1:D20").success, true);
    assert.strictEqual(rangeSchema.safeParse("A1").success, false);
  });

  test("C-08 valuesSchema", () => {
    assert.strictEqual(valuesSchema.safeParse([]).success, false);
    assert.strictEqual(valuesSchema.safeParse([[]]).success, true);
    assert.strictEqual(valuesSchema.safeParse([[null]]).success, true);
    assert.strictEqual(valuesSchema.safeParse([[1, "a", true]]).success, true);
  });

  test("C-09 schema 带中文 describe", () => {
    const schemas = [filePathSchema, sheetNameSchema, cellSchema, rangeSchema, valuesSchema];
    for (const s of schemas) {
      assert.ok(s.description && s.description.length > 0, "schema 应有 description");
    }
  });
});
