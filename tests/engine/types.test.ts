import { test, describe } from "node:test";
import assert from "node:assert";
import type { CellValue, UnifiedSheet, UnifiedWorkbook } from "../../src/engine/types.js";

import * as types from "../../src/engine/types.js";

// 仅做编译期类型检查，运行期确认该文件无导出函数
describe("src/engine/types.ts", () => {
  test("T-01 纯类型文件无运行时代码", () => {
    assert.strictEqual(Object.keys(types).length, 0);
  });

  test("T-01 类型约束可正常使用", () => {
    const value: CellValue = "test";
    assert.strictEqual(value, "test");
    const sheet: UnifiedSheet = {
      name: "Sheet1",
      rowCount: 0,
      colCount: 0,
      values: [],
    };
    const wb: UnifiedWorkbook = { format: "xlsx", sheets: [sheet] };
    assert.strictEqual(wb.format, "xlsx");
  });
});
