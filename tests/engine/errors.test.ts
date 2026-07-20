import { test, describe } from "node:test";
import assert from "node:assert";
import { z } from "zod";
import { ToolError, toUserMessage } from "../../src/engine/errors.js";

describe("src/engine/errors.ts", () => {
  test("E-01 ToolError 携带 code 与 message", () => {
    const e = new ToolError("FILE_NOT_FOUND", "文件不存在：x");
    assert.strictEqual(e.code, "FILE_NOT_FOUND");
    assert.strictEqual(e.message, "文件不存在：x");
  });

  test("E-02 toUserMessage 处理 ToolError", () => {
    const e = new ToolError("SHEET_EXISTS", "工作表已存在");
    assert.strictEqual(toUserMessage(e), "工作表已存在");
  });

  test("E-03 处理 ZodError", () => {
    const schema = z.string();
    let zodError: z.ZodError | undefined;
    try {
      schema.parse(123);
    } catch (e) {
      zodError = e as z.ZodError;
    }
    assert.ok(zodError);
    const msg = toUserMessage(zodError!);
    assert.ok(msg.startsWith("参数错误："));
  });

  test("E-04 处理 ENOENT", () => {
    const e = Object.assign(new Error("x"), { code: "ENOENT", path: "p" });
    const msg = toUserMessage(e);
    assert.ok(msg.includes("文件或目录不存在"));
    assert.ok(msg.includes("p"));
  });

  test("E-05 处理占用类错误", () => {
    for (const code of ["EPERM", "EACCES", "EBUSY"]) {
      const e = Object.assign(new Error("x"), { code });
      assert.strictEqual(
        toUserMessage(e),
        "文件正被其他程序（如 Excel）占用，请关闭后重试。"
      );
    }
  });

  test("E-06 处理普通 Error", () => {
    const e = new Error("boom");
    assert.strictEqual(toUserMessage(e), "操作失败：boom");
  });

  test("E-07 处理非 Error", () => {
    for (const v of ["string", 123, null, { a: 1 }]) {
      assert.strictEqual(toUserMessage(v), "操作失败：未知错误");
    }
  });

  test("E-08 文案不含堆栈", () => {
    const samples = [
      new ToolError("UNKNOWN", "x"),
      Object.assign(new Error("x"), { code: "ENOENT", path: "p" }),
      new Error("boom"),
      "string",
      123,
      null,
      { a: 1 },
    ];
    for (const s of samples) {
      const msg = toUserMessage(s);
      assert.ok(!msg.includes("\n    at "), `堆栈不应出现在：${msg}`);
    }
  });
});
