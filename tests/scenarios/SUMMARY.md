# 多模块联动测试场景 · 汇总报告

> 生成日期：2026-07-20
> 测试依据：`docs/04-test-plan.md` 第 2 章（10 个多模块联动测试场景 S1~S10）
> 总测试数：216（含 23 个测试套件），全部通过 ✅

---

## 一、逐场景测试结果

| 场景 | 测试文件 | 测试数 | 结果 |
|------|---------|:---:|:---:|
| S1 读表问答 | `tests/scenarios/S1-read-chat.test.ts` | 6 | ✅ 全部通过 |
| S2 新建12张考勤表 | `tests/scenarios/S2-create-attendance.test.ts` | 5 | ✅ 全部通过 |
| S3 改表并加合计行 | `tests/scenarios/S3-edit-expense.test.ts` | 3 | ✅ 全部通过 |
| S4 商品清单插图 | `tests/scenarios/S4-product-images.test.ts` | 4 | ✅ 全部通过 |
| S5 数据加工流水线 | `tests/scenarios/S5-data-pipeline.test.ts` | 6 | ✅ 全部通过 |
| S6 格式支持矩阵 | `tests/scenarios/S6-format-matrix.test.ts` | 10 | ✅ 全部通过 |
| S7 原子写故障注入 | `tests/scenarios/S7-atomic-write.test.ts` | 4 | ✅ 全部通过 |
| S8 统一错误链路 | `tests/scenarios/S8-error-chain.test.ts` | 5 | ✅ 全部通过 |
| S9 MCP 端到端冒烟 | `tests/scenarios/S9-e2e-smoke.test.ts` | 2 | ✅ 全部通过 |
| S10 无状态与连续调用 | `tests/scenarios/S10-stateless.test.ts` | 4 | ✅ 全部通过 |
| **合计** | **10 个文件** | **49** | **全部通过 ✅** |

**已有测试回归**：L1（引擎单元测试）、L2（引擎集成测试）、L3（工具集成测试）共 167 个测试全部保持不变，0 个回归失败。

---

## 二、修复的源码 Bug

### Bug 1：`filter_range` 中 `null` 值参与有序比较导致错误匹配

- **文件**：[src/tools/data.ts](src/tools/data.ts) 第 350-359 行
- **问题**：`compareForRelOp()` 函数在比较两个值时，如果其中一个是 `null`，会退化为字符串比较。因为 `String(null) === "null"`，而 `"null".localeCompare("1000")` 返回正数（`'n' > '1'`），导致 `null > 1000` 判定为 `true`。
- **影响**：在 `dedupe_range → sort_range → filter_range` 流水线中，去重后尾部被置为 `null` 的行会被筛选器的数值条件（如 `> 1000`）错误地匹配，导致筛选结果多出不应匹配的行。
- **修复**：在 `compareForRelOp` 函数开头添加 `null` 守卫——如果任一参数为 `null`，返回 `Number.NaN`。由于 JavaScript 中 `NaN` 与任何值的比较（>、>=、<、<=）都返回 `false`，`null` 不再被任何有序操作符匹配。

**修改内容**：
```diff
 function compareForRelOp(a: CellValue, b: CellValue): number {
-  if (a !== null && b !== null) {
-    const numA = Number(a);
-    const numB = Number(b);
-    if (Number.isFinite(numA) && Number.isFinite(numB)) {
-      return numA - numB;
-    }
+  // null 不参与有序比较（> / >= / < / <=），避免 String(null)="null" 的误导性字符串比较
+  if (a === null || b === null) return Number.NaN;
+  const numA = Number(a);
+  const numB = Number(b);
+  if (Number.isFinite(numA) && Number.isFinite(numB)) {
+    return numA - numB;
   }
   return String(a).localeCompare(String(b), "zh-Hans-CN");
 }
```

---

## 三、测试中发现的其他问题（测试本身修正）

在生成 S5（数据加工流水线）测试时，初始测试用例有几处数据量与范围不匹配的问题，均已修正：

1. **S5-01**：测试数据实际为 21 行（1 表头 + 20 数据行），初始断言错写为 20 行 → 修正为 21 行、`cellsWritten: 84`
2. **S5-03**：排序的 range `A1:D5` 不包含第 6 行的"钱七" → 扩展为 `A1:D6`
3. **S5-05**：筛选的 range 同上述问题 → 扩展为 `A1:D6`
4. **S5-06**：去重的 range `A1:D8` 不包含第 9 行的重复"王五" → 扩展为 `A1:D9`；`merge_cells` 破坏表头值导致后续 `filter_range` 无法解析列名 → 将 merge 步骤移至 filter 之后

---

## 四、未解决的问题

**无。** 全部 216 个测试通过（含 49 个场景测试 + 167 个已有回归测试），0 个失败，0 个跳过。

唯一需要注意的是 S9（MCP 端到端冒烟）中的 E2E 测试依赖 `dist/` 目录已构建（`npm run build`），这在 CI 环境或首次运行前需要先执行构建。已有测试 `tests/index.test.ts` 在 `before()` 中自动处理了构建。
