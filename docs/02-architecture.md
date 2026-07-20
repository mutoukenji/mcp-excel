# Excel MCP 工具 · 概要设计文档

> 对应需求文档：`docs/01-requirements.md`
> 设计原则：简单够用。单进程、无数据库、无网络依赖，能跑起来解决问题就行。

## 1. 技术选型

| 选择 | 结论 | 为什么选它 |
|------|------|-----------|
| 运行平台 | **Node.js 20+** | 需求要求 `npx` 一键安装，npx 是 Node 生态的原生能力；Windows / Mac / Linux 全平台免编译运行 |
| 开发语言 | **TypeScript** | MCP 官方 SDK 以 TS 版最成熟、文档最全；类型定义能让工具的输入输出描述更清晰，方便 AI 理解 |
| MCP 框架 | **@modelcontextprotocol/sdk**（stdio 传输） | 官方 SDK，stdio 是本地插件的标准通信方式：AI 助手启动本进程，通过标准输入输出传 JSON 消息，不开端口、不联网 |
| Excel 读取 | **SheetJS（xlsx）** | 唯一能纯 JS 读取老版 `.xls` 的库，同时支持 `.xlsx` / `.csv`，负责"什么格式都能打开" |
| Excel 写入/编辑 | **ExcelJS** | 支持样式、图片、公式、合并单元格，负责"改得漂亮"。它不支持 `.xls`，正好与 SheetJS 互补 |
| 参数校验 | **zod** | MCP SDK 的惯用搭配，一行 schema 同时得到校验和给 AI 看的参数说明 |
| 打包发布 | **tsc 编译 + npm 发布**，包内配 `bin` 入口 | 用户执行 `npx <包名>` 即启动服务，无需构建工具链 |

**为什么是两个 Excel 库而不是一个？** 没有一个纯 JS 库能同时做到"读 .xls"和"写带样式/图片的 .xlsx"。分工：**SheetJS 管读（全格式），ExcelJS 管写（xlsx/csv）**，两者都是纯 JS、无原生依赖，安装即所得，符合"全平台、安装极简"的要求。

**明确不做的**：不引入数据库、不引入 HTTP 服务、不做账号体系（需求明确要求数据不出本机、无需注册）。图表生成（需求 3.2）ExcelJS 不支持，列入 P2 再评估；PDF 导出（需求 3.3）依赖 Office/LibreOffice，v1 不做。

## 2. 模块划分

系统是一个单进程程序，内部分 3 层，共 4 个模块：

```
AI 助手（Kimi / Claude 等，MCP Host）
        │  stdio（JSON-RPC 消息）
        ▼
┌──────────────────────────────────────────┐
│ 模块1：MCP 入口层（server）               │
│   启动进程、建立 stdio 通道、注册全部工具、 │
│   统一错误格式                             │
├──────────────────────────────────────────┤
│ 模块2：工具层（tools）                    │
│   每个需求功能对应一个工具函数，            │
│   用 zod 定义参数 schema + 中文描述        │
├──────────────────────────────────────────┤
│ 模块3：工作簿引擎层（workbook engine）    │
│   3a. 读取适配：SheetJS → 统一内存结构     │
│   3b. 编辑适配：ExcelJS → 修改并保存       │
└──────────────────────────────────────────┘
        │
        ▼
本地文件系统（用户的 Excel 文件，唯一持久存储）
```

模块关系：入口层只做协议收发，不懂 Excel；工具层做参数校验和流程编排；引擎层是唯一接触文件的地方，上层不直接依赖 SheetJS/ExcelJS 的 API，以后换库只改这一层。

目录结构（与模块一一对应）：

```
src/
  index.ts        # 模块1：入口，创建 McpServer，注册工具
  tools/          # 模块2：按功能分文件，如 read.ts / write.ts / format.ts
  engine/         # 模块3：openBook() / saveBook() / readRange() 等
```

## 3. 数据设计

**没有数据库，没有服务端存储。** 系统中只有三类数据：

| 数据 | 存哪 | 生命周期 |
|------|------|---------|
| 用户的 Excel 文件 | 用户本地磁盘 | 唯一的持久数据，本工具只读写、不复制不上传 |
| 内存中的工作簿对象 | 进程内存 | 单次工具调用内有效：打开 → 操作 → 保存 → 释放 |
| 工具参数与返回值（JSON） | stdio 消息 | 一次调用一次传递，不落盘 |

核心设计决策：**无状态调用**。每个工具函数接收完整的 `filePath` 参数，调用时打开文件、做完操作立即保存并释放，不在内存里长期持有打开的文件。好处：实现简单、AI 多次调用之间互不干扰、进程崩溃也不会损坏文件。

写文件安全措施：先写到同目录临时文件，成功后重命名替换原文件，避免写到一半出错把原文件写坏。

## 4. 页面设计

**本产品没有任何自己的页面**（需求第 6 节已明确），因此不存在"页面之间跳转"。用户实际接触的三个触点：

1. **npm 包说明页**：一段静态 README，包含安装命令和 AI 助手配置示例，用户照抄即可。
2. **AI 对话窗口**：属于 AI 助手本身，不在本系统内。
3. **成果文件**：任务完成后用户自己打开的 Excel 文件。

代替页面跳转的是用户的使用流程：

```
复制 npx 命令安装/配置 → 在 AI 助手中用大白话提需求
→ AI 自动调用本工具 → 用户打开 Excel 文件看结果
```

## 5. 接口设计（MCP 工具清单）

所有工具通过 MCP 协议暴露，统一约定：

- 输入第一个参数永远是 `filePath`（绝对路径，字符串）。
- 输出统一为 JSON：`成功 { success: true, ...结果 }`，失败 `{ success: false, error: "人话错误描述" }`。

### 核心工具（v1 必做，对应需求 3.1）

| 名称 | 输入 | 输出 |
|------|------|------|
| `create_workbook` | `filePath`；`sheets?: string[]`（初始 Sheet 名） | `success` |
| `list_sheets` | `filePath` | `sheets: [{ name, rowCount, colCount }]` |
| `read_range` | `filePath`；`sheetName?`（默认第一张）；`range?`（如 `"A1:D20"`，默认全部） | `values: any[][]`（二维数组） |
| `write_range` | `filePath`；`sheetName`；`startCell`（如 `"B5"`）；`values: any[][]` | `success`；`cellsWritten: number` |
| `add_sheet` | `filePath`；`sheetName` | `success` |
| `delete_sheet` | `filePath`；`sheetName` | `success` |
| `rename_sheet` | `filePath`；`oldName`；`newName` | `success` |
| `insert_image` | `filePath`；`sheetName`；`imagePath`；`anchorCell`；`width?/height?` | `success` |

说明：读 `.xls` / `.csv` 时引擎层先转成内存结构再读；单个单元格的增删改由 `write_range` 传入 1×1 数组覆盖，不单独设工具。

### 重要工具（v1 尽量做，对应需求 3.2）

| 名称 | 输入 | 输出 |
|------|------|------|
| `set_formula` | `filePath`；`sheetName`；`cell`；`formula`（如 `"SUM(A1:A10)"`） | `success` |
| `format_cells` | `filePath`；`sheetName`；`range`；`style: { font?, fill?, border?, alignment?, numberFormat? }` | `success` |
| `merge_cells` | `filePath`；`sheetName`；`range` | `success` |
| `sort_range` | `filePath`；`sheetName`；`range`；`keyColumn`；`order: "asc"/"desc"` | `success` |
| `filter_range` | `filePath`；`sheetName`；`range`；`conditions: [{ column, op, value }]` | `values: any[][]`（只返回结果，不改文件） |
| `dedupe_range` | `filePath`；`sheetName`；`range`；`keyColumns: string[]` | `success`；`removedCount` |
| `set_dimensions` | `filePath`；`sheetName`；`columns?: [{ column, width }]`；`rows?: [{ row, height }]`（至少传一个） | `success` |

批量写入即 `write_range` 传大数组，不单独设工具。

### 暂缓（对应需求 3.2 图表及 3.3）

`create_chart`（ExcelJS 不支持，需另评估）、`merge_files`、`set_validation`、模板、PDF 导出——接口预留到 P2/P3，v1 不实现，避免过度设计。

## 6. 核心业务流程

### 流程一：读表问答（需求场景一）

用户："看看《2026年销售记录.xlsx》'一月'那张表，总销售额多少？"

1. AI 判断需要读文件，调用 `read_range(filePath, sheetName="一月")`。
2. 入口层收到请求，转给工具层，zod 校验参数。
3. 引擎层用 SheetJS 打开文件（自动识别 xls/xlsx/csv），定位 Sheet 和区域。
4. 数据转成二维数组，以 JSON 返回给 AI。
5. AI 找到"销售额"列求和，用人话回答用户。

### 流程二：改表保存（需求场景三）

用户："把'费用表'B5 的 3000 改成 3500，最后加一行合计。"

1. AI 先调 `read_range` 读"费用表"，确认 B5 位置和最后一行行号。
2. 调 `write_range` 在 B5 写入 `3500`。
3. 调 `write_range` 在末行写入合计（文字 + `set_formula` 写 SUM 公式）。
4. 引擎层用 ExcelJS 打开文件 → 应用全部修改 → 写临时文件 → 原子替换原文件。
5. 返回 `success`，AI 回复用户"已改好"。

### 流程三：新建考勤表（需求场景二）

用户："新建一个 Excel，12 张 Sheet，每张是一个月考勤表，第一行写姓名和 1~31 号。"

1. AI 调 `create_workbook(filePath, sheets=["一月",…,"十二月"])`。
2. 引擎层用 ExcelJS 新建工作簿、批量建好 12 张 Sheet，保存。
3. AI 对每张 Sheet 调一次 `write_range(startCell="A1", values=[["姓名",1,2,…,31]])`。
4. 全部返回成功，AI 告知用户文件位置，用户打开确认。

—— 三个流程共用同一条主干：**AI 发起 → 校验 → 引擎打开文件 → 读/改 → 保存 → JSON 回执**，没有分支架构，这正是本设计"简单够用"的体现。
