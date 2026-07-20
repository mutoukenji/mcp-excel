# Excel MCP 工具 · 测试计划

> 上游文档：`docs/01-requirements.md`（验收依据）、`docs/02-architecture.md`（模块边界）、`docs/03-detailed-design.md`（逐文件规格，测试用例的判定标准以它为准）。
> 本文档结构：第 0 章测试策略 → 第 1 章逐文件测试用例 → 第 2 章多模块联动场景 → 第 3 章验收检查表。

---

## 0. 测试策略

### 0.1 测试层级

| 层级 | 测什么 | 手段 |
|------|--------|------|
| L1 单元测试 | 纯函数：`address.ts`、`errors.ts`、`normalizeCellValue`、`common.ts` 的 schema/包络/包装器 | 直接 import 函数调用，断言返回值或抛出的 `ToolError` |
| L2 引擎集成测试 | `reader.ts` / `writer.ts` 对真实文件的行为 | 对 `tests/fixtures/` 中的样本文件调用引擎函数；写操作一律先复制到临时目录再测 |
| L3 工具集成测试 | 15 个工具的 handler 业务逻辑 | 直接调用工具 handler（`tool()` 包装后的函数），对临时文件操作，断言返回包络 |
| L4 端到端测试 | `src/index.ts` + MCP 协议 | 用 MCP SDK 的 Client 通过 stdio 启动真实进程，走 initialize → tools/list → tools/call 全流程 |

**重要提示**：zod 参数校验由 MCP SDK 在协议层执行，L3 直接调 handler 会绕过 zod。因此 zod schema 在 L1 用 `schema.safeParse()` 单独测，完整校验链路在 L4 覆盖。

### 0.2 工具与环境（建议，实现时可调整）

- 测试框架：**Node 20 内置 `node:test` + `node:assert`**，不新增测试依赖，符合"简单够用"原则；如需直接跑 TS 源码，可仅加 `tsx` 一个 devDependency。
- 测试目录：`tests/`，用例按源文件命名（如 `tests/engine/address.test.ts`）。
- fixture 文件：放 `tests/fixtures/`，用一个生成脚本（ExcelJS/SheetJS）程序化产出，避免提交二进制文件：
  - `sales.xlsx`：两张表——"一月"（4 列：姓名/销售额/日期/备注，含数字、日期、中间空白行、尾部空行）、"二月"（空表）；
  - `legacy.xls`：同内容的老版格式（SheetJS 以 biff8 写出）；
  - `data.csv`：单表 CSV（含中文逗号分隔值）；
  - `corrupt.xlsx`：内容为纯文本的伪装文件；
  - `img.png` / `img.jpg` / `img.gif` / `img.bmp`：小尺寸图片。
- 每个涉及写操作的用例使用独立临时目录（`fs.mkdtemp`），跑完删除，用例之间互不影响。
- 全平台冒烟：Windows / Mac / Linux 各跑一次完整测试（可手动，也可后续接 CI）。

### 0.3 通用判定约定

- 成功断言：`{ success: true, ... }`，且 MCP 层不设 `isError`。
- 失败断言：`{ success: false, error: <中文人话> }`，且 `isError: true`；`error` 文案须符合设计文档 3.4 节模板（含路径、表名等上下文），不得含英文堆栈。
- 修改类断言：除返回值外，必须用 `readWorkbook()` 或 ExcelJS 重新打开文件验证真实落盘结果（roundtrip）。
- 文件损坏类断言：验证"原文件内容字节不变 + 同目录无 `.tmp-*` 残留"。

---

## 1. 逐文件测试用例

### 1.1 `src/engine/address.ts`（纯函数，L1）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| A-01 | `parseCell` 正常解析 | `parseCell("B5")` | `{ row: 5, col: 2 }` |
| A-02 | 列字母不区分大小写 | `parseCell("b5")`、`parseCell("AB10")` | 与 `"B5"` / `{ row: 10, col: 28 }` 相同 |
| A-03 | 非法单元格地址 | `parseCell("5B")`、`parseCell("")`、`parseCell("A")` | 抛 `ToolError`，code 为 `INVALID_CELL`，文案含正确示例 |
| A-04 | 行号为 0（zod 正则拦不住的情况） | `parseCell("A0")` | 抛 `INVALID_CELL` |
| A-05 | `parseRange` 正常解析 | `parseRange("A1:D20")` | start=`{1,1}`，end=`{20,4}` |
| A-06 | 颠倒区域自动交换 | `parseRange("D20:A1")` | 与 `"A1:D20"` 相同，不报错 |
| A-07 | 非法区域 | `parseRange("A1")`、`parseRange("A1:")`、`parseRange("A:B")` | 抛 `INVALID_RANGE` |
| A-08 | `toCellName` / `colToLetter` | `{row:5,col:2}` → `"B5"`；`colToLetter(1/26/27/28)` | `"B5"`；`"A"/"Z"/"AA"/"AB"` |
| A-09 | `letterToCol` | `letterToCol("a")`、`letterToCol("ab")` | `1`、`28`（大小写不敏感） |
| A-10 | 字母↔列号往返一致 | 对 1~702 循环 `colToLetter` 再 `letterToCol` | 全部回到原值 |
| A-11 | `resolveColumn` 列字母在区域内 | token `"C"`，range `A1:D20` | 返回绝对列号 `3` |
| A-12 | 列字母超出区域 | token `"E"`，range `A1:D20` | 抛 `INVALID_PARAMS`，文案指明"列 E 不在区域 A1:D20 内" |
| A-13 | 表头名精确匹配 | token `"销售额"`，headerRow=`["姓名","销售额","日期"]`，range 起始列 1 | 返回 `2` |
| A-14 | 表头名但未传 headerRow | token `"销售额"`，不传 headerRow | 抛 `INVALID_PARAMS`，文案提示设置 `hasHeader: true` |
| A-15 | 表头名不存在 | token `"销量"`，headerRow 如上 | 抛 `INVALID_PARAMS`，文案列出全部表头 |
| A-16 | 区域不从 A 列开始时表头偏移 | range `C1:F10`，headerRow 命中第 2 项 | 返回 `range.start.col + 1 = 4` |

### 1.2 `src/engine/errors.ts`（L1）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| E-01 | `ToolError` 携带 code 与 message | `new ToolError("FILE_NOT_FOUND", "文件不存在：x")` | `e.code === "FILE_NOT_FOUND"`，`e.message` 原样 |
| E-02 | `toUserMessage` 处理 ToolError | 传入 ToolError | 原样返回 message |
| E-03 | 处理 ZodError | 用某 schema `parse` 非法值触发后传入 | 返回 `"参数错误：" + 第一条 issue 的 message` |
| E-04 | 处理 ENOENT | 构造 `Object.assign(new Error("x"), { code: "ENOENT", path: "p" })` | 返回"文件或目录不存在：…" |
| E-05 | 处理占用类错误 | code 分别为 `EPERM` / `EACCES` / `EBUSY` | 均返回"文件正被其他程序（如 Excel）占用，请关闭后重试。" |
| E-06 | 处理普通 Error | `new Error("boom")` | `"操作失败：boom"` |
| E-07 | 处理非 Error | 传入字符串、number、null、对象 | 一律 `"操作失败：未知错误"` |
| E-08 | 文案不含堆栈 | 以上所有分支 | 返回值不含 `\n    at ` 等堆栈痕迹 |

### 1.3 `src/engine/types.ts`

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| T-01 | 纯类型文件无运行时代码 | `tsc` 编译通过；编译产物中无导出函数 | 编译成功；类型约束由 reader/writer 的用例间接覆盖 |

### 1.4 `src/engine/reader.ts`（L2）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| R-01 | 读 xlsx 多表 | `readWorkbook("sales.xlsx")` | `format: "xlsx"`；两张表名、行列数与 fixture 一致 |
| R-02 | 值归一化 | 检查"一月"各单元格 | 字符串/数字/布尔原样；日期为 ISO 字符串（`2026-01-05T00:00:00.000Z` 形态） |
| R-03 | 行补齐与中间空行保留 | 检查各行长度与空白行 | 每行长度 = colCount；中间空白行为全 null 且保留 |
| R-04 | 尾部全空行裁掉 | fixture 末尾制造全空行 | 裁剪后 `rowCount` 不含尾部空行 |
| R-05 | 空表 | 读"二月" | `{ rowCount: 0, colCount: 0, values: [] }` |
| R-06 | 读 xls | `readWorkbook("legacy.xls")` | `format: "xls"`，内容与 xlsx 版一致 |
| R-07 | 读 csv | `readWorkbook("data.csv")` | `format: "csv"`，单工作表，值正确 |
| R-08 | 文件不存在 | 传入不存在路径 | 抛 `FILE_NOT_FOUND`，文案含路径并提示需要绝对路径 |
| R-09 | 不支持的扩展名 | `notes.txt` | 抛 `UNSUPPORTED_FORMAT`，文案"仅支持 .xlsx / .xls / .csv" |
| R-10 | 文件损坏 | `corrupt.xlsx` | 抛 `UNKNOWN`，文案"无法解析文件…可能已损坏" |
| R-11 | 公式格无缓存结果 | fixture 中写一个无缓存公式格 | 读出的值为 `null` |
| R-12 | 富文本/错误值等异类 | 构造含错误值单元格的文件 | 归一化为 `null` |

### 1.5 `src/engine/writer.ts`（L2）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| W-01 | `createWorkbook` 新建多表 | 新建 3 表文件后用 reader 读回 | 文件存在，3 张空表，名字一致 |
| W-02 | 目标已存在 | 对已存在文件再调 `createWorkbook` | 抛 `FILE_EXISTS`，文案含"如需修改/如需重建"指引；原文件字节不变 |
| W-03 | 目录不存在 | filePath 指向不存在的目录 | 抛 `DIR_NOT_FOUND` |
| W-04 | `editWorkbook` 修改并保存 | 打开 xlsx，mutator 写一格，reader 读回 | 新值落盘，其余内容不变 |
| W-05 | 编辑不存在文件 | filePath 不存在 | 抛 `FILE_NOT_FOUND` |
| W-06 | 编辑 .xls | 对 `legacy.xls` 调 `editWorkbook` | 抛 `READ_ONLY_FORMAT`，文案提示另存为 .xlsx |
| W-07 | 编辑 csv | mutator 改值后读回 | csv 唯一工作表内容被更新 |
| W-08 | 不支持格式 | `.txt` 文件 | 抛 `UNSUPPORTED_FORMAT` |
| W-09 | 损坏文件 | 对 `corrupt.xlsx` 调 `editWorkbook` | 抛 `UNKNOWN`（无法解析文案），原文件不变 |
| W-10 | 原子写：mutator 中途抛错 | mutator 先改一格再抛 `ToolError` | 原文件字节不变；同目录无 `.tmp-*` 残留；错误原样上抛 |
| W-11 | 原子写：保存成功后无残留 | 正常完成一次 edit | 目录下不存在 `<原名>.tmp-<pid>` |
| W-12 | 文件被占用 | mock `fs.renameSync` 抛 `EPERM` | 抛 `FILE_BUSY`，文案提示关闭 Excel 后重试 |
| W-13 | mutator 抛非 ToolError | mutator 抛 `new Error("英文库错误")` | 转成 `UNKNOWN`（由上层 `fail()` 变人话），不裸抛 |
| W-14 | `normalizeCellValue` 全映射 | 逐一喂入设计 2.18 映射表的输入 | null/undefined→null；原始类型原样；Date→ISO；richText→拼接文本；hyperlink→text；formula 取 result（无→null）；error→null；其他对象→null |

### 1.6 `src/tools/common.ts`（L1）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| C-01 | `ok()` 成功包络 | `ok({ a: 1 })` | `content[0].text` 解析出 `{ success: true, a: 1 }`，2 空格缩进，无 `isError` |
| C-02 | `fail()` 失败包络 | `fail(new ToolError("SHEET_EXISTS", "工作表 \"四月\" 已存在。"))` | text 解析出 `success: false` + 人话 error；`isError: true` |
| C-03 | `tool()` 成功路径 | 包装一个返回对象的 handler 并调用 | 得到 ok 包络 |
| C-04 | `tool()` 吞异常 | handler 抛 ToolError / 普通 Error / 字符串 | 永远 resolve，均得到 fail 包络，不 reject |
| C-05 | `filePathSchema` | `safeParse("")`、`safeParse("C:\\a.xlsx")` | 空串失败；正常路径通过 |
| C-06 | `sheetNameSchema` 边界 | 31 字、32 字、含 `[` `]` `:` `*` `?` `/` `\` 的名字 | 31 字通过；其余失败 |
| C-07 | `cellSchema` / `rangeSchema` | `"B5"`/`"5B"`；`"A1:D20"`/`"A1"` | 合法通过、非法失败 |
| C-08 | `valuesSchema` | `[[]]`、`[[null]]`、`[[1,"a",true]]`、`[]` | 空外层数组失败；其余通过（`null` 合法） |
| C-09 | schema 带中文 describe | 读 schema 的 description | 每个字段有中文说明（供 AI 阅读） |

### 1.7 `src/tools/index.ts`（L1/L4）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| TI-01 | 注册全部 15 个工具 | 用 mock server（记录 `registerTool` 调用的名字）调 `registerAllTools` | 恰好 15 个，名字与设计第 5 节清单一一对应，无重复 |

### 1.8 `src/tools/workbook.ts`（5 个工具，L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| WB-01 | `create_workbook` 指定表名 | sheets=`["一月","二月","三月"]` | `success: true, sheetsCreated: 3`；`list_sheets` 验证三表存在 |
| WB-02 | 不传 sheets | 只传 filePath | 创建一张 `"Sheet1"` |
| WB-03 | 文件已存在 | 对已存在文件调用 | `FILE_EXISTS` 文案（含路径与下一步建议），原文件字节不变 |
| WB-04 | 扩展名非 .xlsx | filePath 以 `.xls`/`.csv` 结尾 | `INVALID_PARAMS`（"新建文件的扩展名必须是 .xlsx"） |
| WB-05 | sheets 数组内重名 | `["一月","一月"]` | `INVALID_PARAMS` |
| WB-06 | 非法表名 | 名字含 `:` 或超 31 字 | zod 拦截（L4 验证）或 `INVALID_PARAMS` |
| WB-07 | 目录不存在 | 路径指向不存在目录 | `DIR_NOT_FOUND` |
| WB-08 | `list_sheets` 正常 | 对 sales.xlsx | 返回每张表 `{ name, rowCount, colCount }`，与 reader 结果一致 |
| WB-09 | `list_sheets` 各格式 | xls、csv | 均可列出（读操作不受格式限制） |
| WB-10 | `list_sheets` 文件不存在/格式不支持 | 不存在路径、`.txt` | `FILE_NOT_FOUND` / `UNSUPPORTED_FORMAT` |
| WB-11 | `add_sheet` 正常 | xlsx 加"四月" | 成功；`list_sheets` 出现"四月" |
| WB-12 | `add_sheet` 重名 | 加已存在表名 | `SHEET_EXISTS`（"工作表 \"四月\" 已存在。"） |
| WB-13 | `add_sheet` 对 xls / csv | 分别调用 | xls→`READ_ONLY_FORMAT`；csv→`UNSUPPORTED_FORMAT`（"CSV 文件只有一张工作表，不支持新增"） |
| WB-14 | `delete_sheet` 正常 | 两表中删一张 | 成功；列表中消失 |
| WB-15 | `delete_sheet` 表不存在 | 删"草稿" | `SHEET_NOT_FOUND`，文案附现有表名列表 |
| WB-16 | `delete_sheet` 最后一张 | 单表文件删唯一表 | `LAST_SHEET`（"不能删除最后一张工作表。"） |
| WB-17 | `delete_sheet` 对 xls / csv | 分别调用 | `READ_ONLY_FORMAT` / `UNSUPPORTED_FORMAT` |
| WB-18 | `rename_sheet` 正常 | `Sheet1`→`一季度` | 成功；旧名消失、新名出现，表内容不变 |
| WB-19 | `rename_sheet` 旧名不存在 | oldName 打错 | `SHEET_NOT_FOUND` 附表名列表 |
| WB-20 | `rename_sheet` 新名被占用 | newName 为另一张表名 | `SHEET_EXISTS` |
| WB-21 | `rename_sheet` 改成同名 | oldName === newName | 成功（不视为占用） |
| WB-22 | `rename_sheet` 对 xls / csv | 分别调用 | `READ_ONLY_FORMAT` / `UNSUPPORTED_FORMAT` |

### 1.9 `src/tools/read.ts`（L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| RD-01 | 读整张表 | 传 filePath + sheetName="一月"，不传 range | 返回全部 values |
| RD-02 | 默认第一张表 | 不传 sheetName | 返回 sheets[0] 的内容 |
| RD-03 | 指定区域且超出数据范围 | range="A1:B3"，实际只有 2 行数据 | 返回恰 3 行 2 列，第 3 行补 `[null, null]` |
| RD-04 | 区域行列数严格等于请求 | range="C2:D4" | 返回 3 行 2 列，对应原表坐标值 |
| RD-05 | 表名打错 | sheetName="一月份" | `SHEET_NOT_FOUND`，文案附"现有工作表：一月, 二月" |
| RD-06 | 非法区域 | range="A1:" | `INVALID_RANGE`（或 zod 拦截，L4 验证） |
| RD-07 | 空表 | 读"二月" | `success: true, values: []`（不是错误） |
| RD-08 | 读 xls / csv | 对 legacy.xls、data.csv | 均成功读出值 |
| RD-09 | 日期与空单元格表现 | 检查返回值 | 日期为 ISO 字符串；空格为 null |

### 1.10 `src/tools/write.ts`（2 个工具，L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| WR-01 | 单格写入 | startCell="B5"，values=`[[3500]]` | `cellsWritten: 1`；读回 B5 为 3500，其余格不变 |
| WR-02 | 批量写入 | 写 2 行 4 列到 A1 | `cellsWritten: 8`；读回全部一致 |
| WR-03 | 各行长度不一致 | 两行分别为 3 列、1 列 | 按实际长度写入，`cellsWritten` 按实际格数统计 |
| WR-04 | null 清空单元格 | 对有值的格写 `[[null]]` | 读回为 null；`cellsWritten` 计入该格 |
| WR-05 | `"=..."` 字符串不当公式 | 写 `[["=1+1"]]` | 读回为字符串，不是公式 |
| WR-06 | 表不存在 | sheetName 打错 | `SHEET_NOT_FOUND` |
| WR-07 | 非法 startCell | `"A0"`（zod 拦不住的行号 0） | `INVALID_CELL` |
| WR-08 | 对 xls 写入 | legacy.xls | `READ_ONLY_FORMAT` 文案（提示另存为 xlsx），文件不变 |
| WR-09 | csv 忽略 sheetName | 任意 sheetName 写 data.csv | 写入唯一工作表，读回验证 |
| WR-10 | 文件不存在 | 不存在路径 | `FILE_NOT_FOUND` |
| WR-11 | 文件被占用 | mock renameSync 抛 EPERM（同 W-12 机制走工具层） | 返回 `FILE_BUSY` 人话包络，`isError: true` |
| WR-12 | `set_formula` 不带等号 | formula="SUM(B2:B9)" 写到 B10 | 成功；用 ExcelJS 验证该格 `value.formula === "SUM(B2:B9)"` |
| WR-13 | `set_formula` 带等号 | formula="=AVERAGE(B2:B9)" | 去掉 `=` 后存入，效果同上 |
| WR-14 | 公式去 `=` 后为空 | formula="=" | `INVALID_PARAMS` |
| WR-15 | `set_formula` 已知限制 | 写入后用 reader 读该格 | 值为 null（无缓存结果），README 已声明此行为 |
| WR-16 | `set_formula` 对 xls | legacy.xls | `READ_ONLY_FORMAT` |

### 1.11 `src/tools/format.ts`（3 个工具，L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| F-01 | 字体设置 | style.font={ bold, size, color:"FF0000" } | 成功；ExcelJS 验证 font.bold/size 生效，color.argb 为 `"FFFF0000"` |
| F-02 | 颜色带 `#` 前缀 | color="#305496" | argb 为 `"FF305496"`（去 `#`、转大写、补 alpha） |
| F-03 | 填充色 | fill.color="305496" | cell.fill 为 solid 且 fgColor 正确 |
| F-04 | 边框四边统一 | border={ style:"thin", color:"999999" } | top/left/bottom/right 四边均有相同样式 |
| F-05 | 对齐与自动换行 | alignment={ horizontal:"center", vertical:"middle", wrapText:true } | cell.alignment 一致 |
| F-06 | 数字格式 | numberFormat="0.00" | cell.numFmt === "0.00" |
| F-07 | 只传部分项 | 先设 font，再只设 fill | 第二次调用后 font 保持、fill 更新（未传项不动） |
| F-08 | style 为空对象 | style={} | `INVALID_PARAMS`（"style 至少要包含一项设置"） |
| F-09 | 区域内每格都生效 | range="A1:D1" 设加粗 | 4 个格全部加粗 |
| F-10 | 对 csv / xls | 分别调用 | csv→`UNSUPPORTED_FORMAT`（"CSV 文件不支持样式设置。"）；xls→`READ_ONLY_FORMAT` |
| F-11 | `merge_cells` 正常 | range="A1:D1"，左上角有值 | 成功；ExcelJS 验证合并区域存在，左上角值保留 |
| F-12 | 单格区域 | range="A1:A1" | `INVALID_PARAMS`（必须跨多格） |
| F-13 | 颠倒区域 | range="D1:A1" | 自动交换后合并 A1:D1，不报错 |
| F-14 | 与已有合并区域重叠 | 对重叠区域再次 merge | 返回 fail 包络（`UNKNOWN` 人话），进程不崩 |
| F-15 | `merge_cells` 对 csv | data.csv | `UNSUPPORTED_FORMAT` |
| F-16 | `set_dimensions` 设置列宽 | columns=`[{column:"B",width:16}]` | 成功；ExcelJS 验证 `getColumn(2).width === 16` |
| F-17 | `set_dimensions` 设置行高 | rows=`[{row:1,height:24}]` | 成功；ExcelJS 验证 `getRow(1).height === 24` |
| F-18 | 列宽行高同时设置、列字母小写 | columns=`[{column:"c",width:20}]` + rows=`[{row:2,height:30}]` | 第 3 列宽 20、第 2 行高 30（列字母不区分大小写） |
| F-19 | columns / rows 都未传 | 只传 filePath/sheetName | `INVALID_PARAMS`（"columns 和 rows 至少要传一个"） |
| F-20 | `set_dimensions` 表不存在 | sheetName 打错 | `SHEET_NOT_FOUND` |
| F-21 | `set_dimensions` 对 csv / xls | 分别调用 | csv→`UNSUPPORTED_FORMAT`（"CSV 文件不支持设置列宽行高。"）；xls→`READ_ONLY_FORMAT` |
| F-22 | 设置尺寸不影响单元格内容 | 设置后 `read_range` 读全表 | 各单元格值逐格不变 |

### 1.12 `src/tools/data.ts`（3 个工具，L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| D-01 | 按列字母升序排数字 | keyColumn="B"，order="asc"，无表头 | B 列升序；整行随行移动（行内数据不错位） |
| D-02 | 按表头名降序 | hasHeader=true，keyColumn="销售额"，order="desc" | 表头行不动仍在第一行；数据行按销售额降序 |
| D-03 | null 永远排最后 | 键列含空值，分别 asc/desc | 两种顺序下空值行都在末尾 |
| D-04 | 稳定排序 | 键列有重复值 | 相等行保持原有相对顺序 |
| D-05 | 字符串按中文排序 | 键列为中文姓名 | 按 `localeCompare("zh-Hans-CN")` 排序 |
| D-06 | 数字与字符串混排 | 键列混合类型 | 两个 number 比数值；其余走字符串序，不抛错 |
| D-07 | 列字母超出区域 | keyColumn="E"，range="A1:D20" | `INVALID_PARAMS`（列不在区域内） |
| D-08 | 表头名打错 | keyColumn="销量" | `INVALID_PARAMS`，文案附全部表头 |
| D-09 | 用表头名但 hasHeader 未开 | 传表头名、不传 hasHeader | `INVALID_PARAMS`，提示设置 hasHeader |
| D-10 | 区域外内容不动 | 排序后读全表 | 区域外单元格逐格不变 |
| D-11 | csv 排序可用 | 对 data.csv 排序 | 成功（值操作不拦 csv） |
| D-12 | 对 xls 排序 | legacy.xls | `READ_ONLY_FORMAT` |
| D-13 | `filter_range` 数值比较 | conditions=[{ column:"销售额", op:">", value:1000 }]，hasHeader | 只返回满足行，表头在结果第一行 |
| D-14 | 多条件"并且" | 两个条件同时给 | 仅返回同时满足的行 |
| D-15 | 宽松相等 | 单元格为 number 1000，条件 value 为字符串 "1000"，op="=" | 判定相等（类型不等再比 String 形式） |
| D-16 | `!=` / `>=` / `<=` / `<` | 各构造边界值 | 比较结果正确（含等于边界） |
| D-17 | contains 大小写敏感 | 单元格 "Apple"，value "app" | 不匹配 |
| D-18 | 字符串间的大小比较 | 两侧不能转数字时用 op=">" | 走 localeCompare 字符串序，不抛错 |
| D-19 | 0 行匹配 | 条件无人满足，hasHeader=true / false | 前者返回仅表头行，后者返回 `[]`；均 `success: true` |
| D-20 | 不改文件 | 记录调用前后文件字节 | 完全一致（filter 是只读操作） |
| D-21 | xls 允许筛选 | 对 legacy.xls 调 filter | 成功（只读操作不受 READ_ONLY_FORMAT 限制） |
| D-22 | `dedupe_range` 基本去重 | keyColumns=["A"]，含重复行 | 保留首现行；`removedCount` 正确 |
| D-23 | 多键组合 | keyColumns=["A","C"] | 仅当两列都相同才判重 |
| D-24 | 尾部置空不移动区域外 | 区域下方有数据 | 被删行变全 null；区域外行的位置和值不变 |
| D-25 | hasHeader 参与 | hasHeader=true 且首行与某数据行键相同 | 表头不参与去重、不被删 |
| D-26 | 无重复 | 全部唯一 | `removedCount: 0`，内容不变 |
| D-27 | 归一化后判重 | 一行键为 number 1、另一行为字符串 "1" | 按归一化值 + JSON.stringify 判重，行为与设计一致（类型不同不算重复） |

### 1.13 `src/tools/image.ts`（L3）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| I-01 | 插入 png 显式尺寸 | width=120, height=120，anchorCell="D2" | 成功；ExcelJS 验证该表有 1 张图片，tl 锚点为 D2（内部 0 起始转换正确）、尺寸 120×120 |
| I-02 | 默认尺寸 | 不传 width/height | 图片尺寸为 300×200 |
| I-03 | 扩展名映射 | 分别插入 .jpg / .jpeg / .gif | 均成功（jpg、jpeg 都映射为 jpeg） |
| I-04 | 不支持的图片格式 | img.bmp | `UNSUPPORTED_IMAGE`（"仅支持 png / jpg / gif"） |
| I-05 | 图片不存在 | imagePath 指向不存在文件 | `IMAGE_NOT_FOUND`，文案含图片路径 |
| I-06 | 锚点非法 | anchorCell="A0" | `INVALID_CELL` |
| I-07 | 表不存在 | sheetName 打错 | `SHEET_NOT_FOUND` |
| I-08 | 对 csv / xls | 分别调用 | csv→`UNSUPPORTED_FORMAT`；xls→`READ_ONLY_FORMAT` |
| I-09 | 锚点 0 起始换算 | anchorCell="A1" | 图片 tl 为 `{ col: 0, row: 0 }`（不错位到 B2） |

### 1.14 `src/index.ts`（L4 + 静态检查）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| IX-01 | shebang 保留 | 读 `src/index.ts` 第一行与编译后 `dist/index.js` 第一行 | 均为 `#!/usr/bin/env node` |
| IX-02 | stdout 纯净 | 全项目 grep `console.log` / `process.stdout.write`；E2E 抓 stdout | 源码 0 处命中；运行期 stdout 每一行都是合法 JSON-RPC 消息 |
| IX-03 | 进程启动并待命 | stdio 启动进程，发送 initialize | 握手成功，进程保持运行不退出 |
| IX-04 | 启动日志走 stderr | 抓 stderr | 有"已启动"提示；stdout 无该内容 |
| IX-05 | 启动失败退出码 | 制造 connect 失败（如抢占 stdin，可选） | 打印 stderr 并以退出码 1 结束 |

### 1.15 配置 / 发布文件（L4 + 静态检查）

| 编号 | 测什么 | 怎么测 | 预期结果 |
|------|--------|--------|---------|
| CF-01 | 构建通过 | `npm install && npm run build` | `tsc` 无错误，产出 `dist/` |
| CF-02 | bin 入口可用 | `node dist/index.js`（及发布前 `npm pack` 后本地 `npx` 试装） | 进程启动等待 stdio |
| CF-03 | 发布内容最小化 | `npm pack --dry-run` 看文件清单 | 只含 `dist/`（及 README/package.json 等 npm 默认文件） |
| CF-04 | engines 声明 | 读 package.json | `node >= 20` |
| CF-05 | .gitignore 生效 | 构建后 `git status --porcelain` | `dist/`、`node_modules/` 不出现在变更列表 |
| CF-06 | README 配置示例可用 | 提取 README 中 MCP 配置 JSON 块解析校验 | 是合法 JSON；含 `npx -y mcp-excel`；含 Node 20 要求说明 |
| CF-07 | README 如实声明限制 | 通读 README | 写明格式支持矩阵、已知限制（图表/透视表/宏丢失、公式需打开计算）、隐私说明 |

---

## 2. 多模块联动测试场景

> 每个场景走"工具层 → 引擎层 → 真实文件 → 再读回验证"的完整链路，覆盖单一文件用例测不到的协作关系。场景 S1~S4 直接对应需求文档第 5 节的四个典型场景。

### S1 读表问答（需求场景一 / 架构流程一）

- **涉及模块**：`tools/read.ts`、`tools/workbook.ts`(list_sheets)、`engine/reader.ts`、`engine/address.ts`
- **步骤**：
  1. 对 `sales.xlsx` 调 `list_sheets` 确认"一月"存在；
  2. 调 `read_range(filePath, sheetName="一月")`；
  3. 模拟 AI 行为：在返回的二维数组中定位"销售额"列并求和。
- **预期**：两步均 `success: true`；返回值结构与设计 3.1 节示例一致（日期为 ISO 字符串、空格为 null）；求和结果与 fixture 预置总额一致。

### S2 新建 12 张考勤表（需求场景二 / 架构流程三）

- **涉及模块**：`tools/workbook.ts`、`tools/write.ts`、`engine/writer.ts`、`engine/reader.ts`
- **步骤**：
  1. `create_workbook(filePath, sheets=["一月",…,"十二月"])`；
  2. 对每张表各调一次 `write_range(startCell="A1", values=[["姓名",1,2,…,31]])`；
  3. `list_sheets` + 抽查 3 张表 `read_range("A1:AF1")` 验证。
- **预期**：`sheetsCreated: 12`；12 次写入全部成功；每张表首行恰 32 列且内容正确；各次调用互不干扰（验证无状态设计）。

### S3 改表并加合计行（需求场景三 / 架构流程二）

- **涉及模块**：`tools/read.ts`、`tools/write.ts`、`engine/writer.ts`、`engine/address.ts`
- **步骤**：
  1. 准备含"费用表"的 xlsx（B5=3000，共 9 行）；
  2. `read_range` 确认 B5 位置与末行行号；
  3. `write_range(startCell="B5", values=[[3500]])`；
  4. `write_range` 在第 10 行写 `["合计", null]`；
  5. `set_formula(cell="B10", formula="SUM(B2:B9)")`；
  6. 读回全表验证。
- **预期**：B5=3500；A10="合计"；ExcelJS 验证 B10 公式为 `SUM(B2:B9)`；其余单元格不变；用 Excel/WPS 打开（手动验收）公式显示计算结果。

### S4 商品清单插图（需求场景四）

- **涉及模块**：`tools/workbook.ts`、`tools/write.ts`、`tools/image.ts`、`engine/writer.ts`
- **步骤**：
  1. `create_workbook` 新建"商品清单.xlsx"（含"清单"表）；
  2. `write_range` 写入 3 行商品数据；
  3. 对每个商品 `insert_image(anchorCell="D2"/"D3"/"D4", width=120, height=120)` 插 3 张图；
  4. ExcelJS 读回验证。
- **预期**：3 张图分别锚定 D2/D3/D4、尺寸正确；单元格数据不受影响；文件可被 Excel 正常打开（手动验收）。

### S5 数据加工流水线（write → format → merge → sort → dedupe → filter）

- **涉及模块**：6 个工具文件 + 全部引擎文件（覆盖工具间对同一文件的接力修改）
- **步骤**：
  1. `create_workbook` + `write_range` 写入 20 行含重复、未排序、无样式的数据；
  2. `dedupe_range(keyColumns=["A"], hasHeader=true)`；
  3. `sort_range(keyColumn="销售额", order="desc", hasHeader=true)`；
  4. `format_cells` 美化标题行；`merge_cells("A1:D1")` 合并标题（若标题独立于表头，先 insert 行再合并，按实际 fixture 设计）；
  5. `filter_range(销售额 > 1000)` 取结果；
  6. `read_range` 读全表终态验证。
- **预期**：每一步都基于上一步的落盘结果正确执行；终态文件 = 已去重 + 已排序 + 有样式 + 有合并；filter 结果与终态文件内容一致（证明读写两条引擎路径数据口径一致）。

### S6 格式支持矩阵全量核对（全局约定 8）

- **涉及模块**：全部工具 × reader/writer 的格式分支
- **步骤**：对同一份内容的 `.xlsx` / `.xls` / `.csv` 三个副本，分别执行 15 个工具中的代表性操作（读：list_sheets/read_range/filter_range；写值：write_range/set_formula/sort/dedupe；表操作：add/delete/rename；样式类：format/merge/insert_image/set_dimensions）。
- **预期**：每个（操作 × 格式）组合的结果与设计 0.8 节矩阵**逐格一致**：xlsx 全 ✅；xls 读 ✅、写一律 `READ_ONLY_FORMAT`；csv 读写值 ✅（sort/dedupe 可用）、表操作与样式类一律 `UNSUPPORTED_FORMAT` 且文案为 CSV 专用文案。

### S7 原子写故障注入（写坏保护链路）

- **涉及模块**：`engine/writer.ts`、`engine/errors.ts`、`tools/common.ts`
- **步骤**：
  1. 对有内容的 xlsx 调 `write_range`，用 mock 让 `renameSync` 抛 `EPERM`；
  2. 另做一组：让 ExcelJS `writeFile` 抛 `ENOENT`（目录不存在场景）；
  3. 检查原文件字节、临时文件残留、返回包络。
- **预期**：两种故障下原文件字节均不变；无 `.tmp-*` 残留；分别返回 `FILE_BUSY` / `DIR_NOT_FOUND` 人话包络且 `isError: true`。

### S8 统一错误链路（三条错误路径殊途同归）

- **涉及模块**：`tools/common.ts`、`engine/errors.ts`、zod、各工具
- **步骤**（L4 端到端下验证）：
  1. zod 路径：传非法 sheetName（含 `:`）调 `add_sheet`；
  2. ToolError 路径：读不存在的表；
  3. 未知异常路径：对 `corrupt.xlsx` 调 `read_range`。
- **预期**：三者都返回 `{ success: false, error: <中文人话> }` 且 `isError: true`；error 文案分别命中"参数错误：…"、`SHEET_NOT_FOUND` 模板、"无法解析文件…"模板；协议层看不到异常堆栈。

### S9 MCP 端到端冒烟（模块 1 + 全量注册）

- **涉及模块**：`src/index.ts`、`tools/index.ts`、全部工具
- **步骤**：用 MCP Client 经 stdio 启动真实进程 → initialize → `tools/list` → `tools/call` 依次调 `create_workbook`、`write_range`、`read_range` → 关闭进程。
- **预期**：`tools/list` 返回恰好 15 个工具，每个都有中文 title/description 和参数 describe；三次 call 结果包络正确；全程 stdout 只出现 JSON-RPC 消息（无杂散输出污染协议）；进程可正常退出。

### S10 无状态与连续调用（全局约定 1）

- **涉及模块**：`engine/writer.ts` 与全部写类工具
- **步骤**：对同一文件快速连续调用 20 次 `write_range`（不同单元格），随后 `read_range` 全量读回；再对两个不同文件交替调用写操作。
- **预期**：20 次写入全部生效、无丢失无串扰（每次调用独立打开/保存/释放）；两文件内容各自正确。

---

## 3. 验收检查表

> 判定方法：逐条执行后勾选「通过」或「不通过」。全部核心项（3.1）通过才算 v1 验收通过；标注【暂缓】的条目确认"v1 未实现且文档如实说明"即视为通过。

### 3.1 核心功能（需求 3.1，必须全部通过）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 1 | 给定绝对路径，能读出任意 xlsx 文件的内容 | `read_range` 返回正确二维数组（联动 S1） | □ 通过 / □ 不通过 |
| 2 | 能从无到有创建新的 Excel 文件 | `create_workbook` 生成可用文件，支持指定初始 Sheet（S2） | □ 通过 / □ 不通过 |
| 3 | 能对任意 Sheet 增、删、改 | `add_sheet` / `delete_sheet` / `rename_sheet` / `write_range` 单格修改全部生效（WB-11~22、WR-01、S3） | □ 通过 / □ 不通过 |
| 4 | `.xlsx`、`.xls`、`.csv` 都能读取 | 三种格式 `read_range` 均正确（RD-08、S6） | □ 通过 / □ 不通过 |
| 5 | 格式限制有明确中文报错而非乱码崩溃 | xls 写操作、csv 样式/表操作均返回人话错误（S6） | □ 通过 / □ 不通过 |
| 6 | 能把本地图片插入表格指定位置 | png/jpg/gif 插入成功并锚定正确（I-01~03、S4） | □ 通过 / □ 不通过 |
| 7 | 用户复制一条 npx 命令即可完成安装配置 | 干净环境（无全局依赖）按 README 配置后 AI 助手可调通（CF-02/06、S9） | □ 通过 / □ 不通过 |

### 3.2 重要功能（需求 3.2）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 8 | 能在单元格写入求和等计算公式 | `set_formula` 写入成功，Excel/WPS 打开后显示计算结果（WR-12/13、S3） | □ 通过 / □ 不通过 |
| 9 | 能设置字体、颜色、边框、对齐、数字格式 | `format_cells` 各项生效（F-01~07） | □ 通过 / □ 不通过 |
| 10 | 能合并单元格 | `merge_cells` 生效且保留左上角值（F-11） | □ 通过 / □ 不通过 |
| 11 | 能排序、按条件筛选、去重 | 三个工具行为符合设计（D-01~27、S5） | □ 通过 / □ 不通过 |
| 12 | 能一次写入大片数据 | `write_range` 多行多列一次落盘（WR-02、S2 的 12 表写入） | □ 通过 / □ 不通过 |
| 13 | 能调整列宽和行高 | `set_dimensions` 设置列宽/行高生效（F-16~F-18），在 Excel/WPS 中打开确认显示正常 | □ 通过 / □ 不通过 |
| 14 | 【暂缓】生成图表 | v1 明确不实现（ExcelJS 不支持，P2 再评估）；确认 README/文档未虚假声称支持即通过 | □ 通过 / □ 不通过 |

### 3.3 锦上添花功能（需求 3.3，v1 全部暂缓）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 15 | 【暂缓】多文件合并 | v1 不实现，文档如实说明 | □ 通过 / □ 不通过 |
| 16 | 【暂缓】数据校验 | v1 不实现，文档如实说明 | □ 通过 / □ 不通过 |
| 17 | 【暂缓】模板能力 | v1 不实现，文档如实说明 | □ 通过 / □ 不通过 |
| 18 | 【暂缓】导出 PDF | v1 不实现，文档如实说明 | □ 通过 / □ 不通过 |

### 3.4 额外要求（需求第 4 节）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 19 | Windows / Mac / Linux 三平台都能安装运行 | 三平台各跑一次 L1~L4 测试全绿 + S9 冒烟通过 | □ 通过 / □ 不通过 |
| 20 | 安装无需构建工具链和复杂环境 | 仅需 Node 20+，无原生编译依赖（xlsx/exceljs 均为纯 JS） | □ 通过 / □ 不通过 |
| 21 | 数据不出本机 | 代码审查无网络调用；运行时不监听端口、不发起外网请求（可用 netstat/抓包验证） | □ 通过 / □ 不通过 |
| 22 | 无需注册、登录、无管理后台 | 代码中无账号/鉴权逻辑 | □ 通过 / □ 不通过 |
| 23 | 每个工具都有让 AI 一看就懂的中文描述 | `tools/list` 返回的 15 个工具均有中文 title/description，每个参数有中文 describe（C-09、S9） | □ 通过 / □ 不通过 |

### 3.5 健壮性与设计约束（来自架构/详细设计，属于隐性验收项）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 24 | 写操作失败时原文件绝不损坏 | S7 两种故障注入后原文件字节不变 | □ 通过 / □ 不通过 |
| 25 | 任何情况下不在用户目录留下临时文件垃圾 | S7 + W-11 验证无 `.tmp-*` 残留 | □ 通过 / □ 不通过 |
| 26 | 多次调用互不干扰（无状态） | S10 连续/交替调用结果全部正确 | □ 通过 / □ 不通过 |
| 27 | stdout 没有任何协议外输出 | IX-02 静态 grep + S9 运行时验证 | □ 通过 / □ 不通过 |
| 28 | 所有错误都是中文人话、给出下一步建议、不暴露堆栈 | S8 三条错误路径 + E-08 | □ 通过 / □ 不通过 |
| 29 | "没数据"不报错：空表/筛选无匹配返回 `success: true` | RD-07、D-19 | □ 通过 / □ 不通过 |
| 30 | README 如实写明已知限制（图表/透视表/宏可能丢失、公式需打开一次才显示结果） | CF-07 人工核对 | □ 通过 / □ 不通过 |

### 3.6 四个典型场景端到端验收（需求第 5 节）

| # | 检查项（判断句） | 通过标准 | 结果 |
|---|----------------|---------|------|
| 31 | 场景一"看表"可完成：读出'一月'表并让 AI 算出总销售额 | S1 通过 | □ 通过 / □ 不通过 |
| 32 | 场景二"建表"可完成：一次建好 12 张考勤 Sheet 并填好表头 | S2 通过 | □ 通过 / □ 不通过 |
| 33 | 场景三"改表"可完成：B5 改为 3500 并追加合计公式行 | S3 通过 | □ 通过 / □ 不通过 |
| 34 | 场景四"插图"可完成：产品图插到对应商品旁 | S4 通过 | □ 通过 / □ 不通过 |

---

## 附：遗留风险与说明

- **图表功能**：ExcelJS 不支持，架构文档已明确列入 P2，v1 验收不含此项。
- **手动验收项**：文件在真实 Excel/WPS 中打开的观感（样式、图片、公式计算结果）无法完全自动化，S3/S4 的最后一步需人工确认。
- **fixture 的二进制依赖**：`.xls` fixture 依赖 SheetJS 的 biff8 写出能力生成；若生成文件与真实 Office 产出的 xls 有差异，建议补充一个真实 Excel 保存的样本做对照测试。
