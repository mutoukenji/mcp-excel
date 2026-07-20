// 对应 docs/03-detailed-design.md 第 2.7 节：src/tools/common.ts 工具层共享地基

import { z } from "zod";
import { toUserMessage } from "../engine/errors.js";

// 对应第 2.7 节：MCP 文本结果包络类型（handler 返回形态）
export type McpTextResult = {
  content: Array<{ type: "text"; text: string }>;
  isError?: boolean;
};

// 对应第 2.7 节 (1)：zod schema，每个都带 .describe() 中文说明

/** 对应第 2.7 节：Excel 文件的绝对路径 */
export const filePathSchema = z
  .string()
  .min(1)
  .describe(
    "Excel 文件的绝对路径，例如 `C:\\Users\\xx\\报表.xlsx` 或 `/Users/xx/报表.xlsx`。支持 .xlsx / .xls / .csv"
  );

/** 对应第 2.7 节：工作表名称，1~31 个字符，不能包含 [ ] : * ? / \\ */
export const sheetNameSchema = z
  .string()
  .min(1)
  .max(31)
  .regex(/^[^\[\]\:\*\?\/\\]+$/)
  .describe("工作表名称，1~31 个字符，不能包含 [ ] : * ? / \\");

/** 对应第 2.7 节：单元格地址，如 `B5` */
export const cellSchema = z
  .string()
  .regex(/^[A-Za-z]+[0-9]+$/)
  .describe("单元格地址，如 `B5`");

/** 对应第 2.7 节：矩形区域，如 `A1:D20` */
export const rangeSchema = z
  .string()
  .regex(/^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/)
  .describe("矩形区域，如 `A1:D20`");

/** 对应第 2.7 节：二维数组，每个内层数组是一行。`null` 表示清空该单元格 */
export const valuesSchema = z
  .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
  .min(1)
  .describe("二维数组，每个内层数组是一行。`null` 表示清空该单元格");

// 对应第 2.7 节 (2)：返回包络

/**
 * 对应第 2.7 节：成功包络
 * 把 { success: true, ...data } 序列化为 2 空格缩进的 JSON 文本
 */
export function ok(data: Record<string, unknown>): McpTextResult {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: true, ...data }, null, 2),
      },
    ],
  };
}

/**
 * 对应第 2.7 节：失败包络
 * 任何异常 → 经 toUserMessage 转换为人话 → { success: false, error }
 * 同时置 isError: true，通知协议层这是一次失败调用
 */
export function fail(e: unknown): McpTextResult {
  const message = toUserMessage(e);
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: false, error: message }, null, 2),
      },
    ],
    isError: true,
  };
}

// 对应第 2.7 节 (3)：tool() 错误包装器

/**
 * 对应第 2.7 节：tool() 包装器
 * 输入：业务 handler（返回“结果字段对象”，不含 success）
 * 输出：符合 SDK 要求的 handler，内部 try { return ok(await handler(args)) } catch { return fail(e) }
 * 效果：业务 handler 只管抛 ToolError 或返回结果，永远不需要自己 try/catch
 */
export function tool<A extends Record<string, unknown>>(
  handler: (
    args: A
  ) => Promise<Record<string, unknown>> | Record<string, unknown>
): (args: A) => Promise<McpTextResult> {
  return async (args: A): Promise<McpTextResult> => {
    try {
      return ok(await handler(args));
    } catch (e) {
      return fail(e);
    }
  };
}
