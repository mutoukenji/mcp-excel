// 对应 docs/03-detailed-design.md 第 2.15 节：src/engine/errors.ts 错误类型与异常转换

import { z } from "zod";

// 对应第 2.15 节 (1)：错误码
export type ErrorCode =
  | "FILE_NOT_FOUND" | "DIR_NOT_FOUND" | "FILE_EXISTS" | "FILE_BUSY"
  | "UNSUPPORTED_FORMAT" | "READ_ONLY_FORMAT"
  | "SHEET_NOT_FOUND" | "SHEET_EXISTS" | "LAST_SHEET"
  | "INVALID_CELL" | "INVALID_RANGE" | "INVALID_PARAMS"
  | "IMAGE_NOT_FOUND" | "UNSUPPORTED_IMAGE"
  | "UNKNOWN";

/**
 * 对应第 2.15 节 (1)：
 * 全项目统一的业务错误类。
 * message 直接就是给最终用户看的中文文案（含具体路径、表名等上下文）。
 */
export class ToolError extends Error {
  readonly code: ErrorCode;
  constructor(code: ErrorCode, message: string) {
    super(message);
    this.code = code;
  }
}

/**
 * 对应第 2.15 节 (2)：
 * 把任何异常（ToolError、zod 校验错误、NodeJS 文件错误、库抛错）
 * 转换成最终用户/AI 能看到的人话，绝不暴露堆栈。
 */
export function toUserMessage(e: unknown): string {
  // ToolError：message 已经是人话，原样返回
  if (e instanceof ToolError) {
    return e.message;
  }

  // zod 校验错误：取第一条 issue 的 message
  if (e instanceof z.ZodError) {
    const first = e.issues[0];
    return `参数错误：${first ? first.message : "参数校验失败"}`;
  }

  // Node.js 文件系统错误：按 code 给标准文案
  if (e instanceof Error) {
    const nodeErr = e as NodeJS.ErrnoException;
    const code = nodeErr.code;

    if (code === "ENOENT") {
      const pathInfo = nodeErr.path ?? e.message;
      return `文件或目录不存在：${pathInfo}`;
    }

    if (code === "EPERM" || code === "EACCES" || code === "EBUSY") {
      return "文件正被其他程序（如 Excel）占用，请关闭后重试。";
    }

    return `操作失败：${e.message}`;
  }

  // 非 Error 对象
  return "操作失败：未知错误";
}
