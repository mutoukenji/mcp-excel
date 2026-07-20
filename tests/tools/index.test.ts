import { test, describe } from "node:test";
import assert from "node:assert";
import { registerAllTools } from "../../src/tools/index.js";

describe("src/tools/index.ts", () => {
  test("TI-01 注册全部 15 个工具", () => {
    const registered: string[] = [];
    const server = {
      registerTool: (name: string, _meta: unknown, _handler: unknown) => {
        registered.push(name);
      },
    };
    registerAllTools(server as any);
    const expected = [
      "create_workbook",
      "list_sheets",
      "add_sheet",
      "delete_sheet",
      "rename_sheet",
      "read_range",
      "write_range",
      "set_formula",
      "format_cells",
      "merge_cells",
      "set_dimensions",
      "sort_range",
      "filter_range",
      "dedupe_range",
      "insert_image",
    ];
    assert.deepStrictEqual(registered, expected);
    assert.strictEqual(new Set(registered).size, 15);
  });
});
