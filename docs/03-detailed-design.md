# Excel MCP 工具 · 详细设计文档

> 上游文档：`docs/02-architecture.md`（技术选型、模块划分、工具清单以它为准）。
> 本文档自足：编码时只看本文档即可写出全部代码，不需要回看需求文档。
> 文档结构：第 0 章全局约定 → 第 1 章文件清单 → 第 2 章逐文件规格 → 第 3 章数据格式 → 第 4 章触点（"页面"）设计。

---

## 0. 全局约定（写任何一行代码前先记住这些）

**技术栈与参考版本**（安装时取 npm 上最新兼容版本即可）：

| 依赖 | 参考版本 | 用途 |
|------|---------|------|
| node | >= 20 | 运行平台 |
| typescript | ^5.5 | 编译到 `dist/` |
| @modelcontextprotocol/sdk | ^1.x | MCP 服务端（stdio） |
| zod | ^3.25 | 参数 schema（SDK 的 `registerTool` 需要 zod v3 形状） |
| xlsx（SheetJS） | ^0.18.5 | 读取 xls / xlsx / csv |
| exceljs | ^4.4 | 新建 / 修改 / 保存 xlsx 与 csv |

**全局规则：**

1. **无状态调用**：每个工具调用 = 打开文件 → 操作 → 保存 → 释放。不在内存里缓存任何打开的文件。
2. **原子写**：保存一律先写同目录临时文件 `<原名>.tmp-<pid>`，成功后 `fs.renameSync` 替换原文件。任何工具不得直接覆盖写原文件。
3. **stdout 是协议通道**：全项目禁止 `console.log` / `process.stdout.write`（会破坏 MCP 的 JSON-RPC 消息）。日志只走 `console.error`（stderr）。
4. **统一返回包络**：所有工具 handler 永远正常 resolve，不向协议层抛异常。业务结果统一序列化为 JSON 文本：
   - 成功 `{ "success": true, ...结果字段 }`
   - 失败 `{ "success": false, "error": "人话错误描述" }`（同时置 MCP 的 `isError: true`）
5. **统一错误处理**：业务错误抛 `ToolError`（见 2.15），工具层用 `tool()` 包装器统一捕获转成失败包络（见 2.7）。
6. **ESM + NodeNext**：`package.json` 设 `"type": "module"`，TypeScript 源码中相对导入必须带 `.js` 后缀（如 `import { x } from "./engine/types.js"`）。
7. **读用 SheetJS，写用 ExcelJS**：读取走 `engine/reader.ts`（SheetJS，全格式），修改走 `engine/writer.ts`（ExcelJS，仅 xlsx / csv）。上层工具不直接 import 这两个库。
8. **格式支持矩阵**（后续每章错误处理都以此为准）：

| 操作 | .xlsx / .xlsm | .xls | .csv |
|------|:---:|:---:|:---:|
| 读取（list_sheets / read_range / filter_range） | ✅ | ✅ | ✅ |
| 写入值 / 公式（write_range / set_formula / sort / dedupe） | ✅ | ❌ 报 `READ_ONLY_FORMAT` | ✅（仅唯一工作表） |
| 工作表增删改名 | ✅ | ❌ 报 `READ_ONLY_FORMAT` | ❌ 报 `UNSUPPORTED_FORMAT`（csv 单表） |
| 样式 / 合并 / 图片 / 列宽行高 | ✅ | ❌ 报 `READ_ONLY_FORMAT` | ❌ 报 `UNSUPPORTED_FORMAT`（csv 无样式） |

9. **地址体系**：单元格 `"B5"`、区域 `"A1:D20"`，列字母不区分大小写，行列均为 1 起始。只在 `engine/address.ts` 一处做解析。
10. **已知限制**（写进 README，不用代码解决）：用本工具修改含图表、数据透视表、宏（xlsm 的 VBA）的文件时，这些元素会在 ExcelJS 往返保存中丢失；本工具写入的公式不带缓存计算结果，需用户在 Excel/WPS 中打开一次后才会显示计算值。

---

## 1. 源代码文件清单

后续编码严格按此清单创建文件，共 18 个文件。

**配置 / 发布文件（4 个）：**

| 文件路径 | 一句话职责 |
|---------|-----------|
| `package.json` | 包元数据、依赖、`bin` 入口、构建脚本 |
| `tsconfig.json` | TypeScript 编译配置（ESM / NodeNext） |
| `.gitignore` | 忽略 `node_modules/` 与 `dist/` |
| `README.md` | npm 包说明页内容（安装命令 + 配置示例），布局见第 4 章 |

**入口层（1 个）：**

| 文件路径 | 一句话职责 |
|---------|-----------|
| `src/index.ts` | 进程入口：创建 McpServer、注册全部工具、接 stdio 传输 |

**工具层（7 个）：**

| 文件路径 | 一句话职责 |
|---------|-----------|
| `src/tools/index.ts` | `registerAllTools()`：汇总调用各分文件的注册函数 |
| `src/tools/common.ts` | 共享 zod schema、成功/失败包络、`tool()` 错误包装器 |
| `src/tools/workbook.ts` | 5 个工具：create_workbook / list_sheets / add_sheet / delete_sheet / rename_sheet |
| `src/tools/read.ts` | 1 个工具：read_range |
| `src/tools/write.ts` | 2 个工具：write_range / set_formula |
| `src/tools/format.ts` | 3 个工具：format_cells / merge_cells / set_dimensions |
| `src/tools/data.ts` | 3 个工具：sort_range / filter_range / dedupe_range |
| `src/tools/image.ts` | 1 个工具：insert_image |

**引擎层（5 个）：**

| 文件路径 | 一句话职责 |
|---------|-----------|
| `src/engine/types.ts` | 统一内存结构类型：`CellValue` / `UnifiedSheet` / `UnifiedWorkbook` |
| `src/engine/errors.ts` | `ToolError` 类、错误码、`toUserMessage()` 异常→人话转换 |
| `src/engine/address.ts` | 单元格/区域/列字母解析，`resolveColumn()` 列标识解析 |
| `src/engine/reader.ts` | `readWorkbook()`：SheetJS 打开任意格式 → `UnifiedWorkbook` |
| `src/engine/writer.ts` | `createWorkbook()` / `editWorkbook()`：ExcelJS 修改并原子保存 |

**文件依赖方向**（只允许自上而下，禁止反向、禁止循环）：

```
src/index.ts → src/tools/* → src/engine/* → (xlsx / exceljs / node 内置模块)
src/tools/common.ts 只依赖 src/engine/errors.ts 与 zod
src/engine/types.ts 不依赖任何项目内文件
```

---

## 2. 每个文件的详细规格

### 2.1 `package.json`

**负责什么**：npm 包定义。用户 `npx mcp-excel` 时，npm 依据 `bin` 字段找到并执行 `dist/index.js`。

**完整内容规格**（字段级）：

```json
{
  "name": "mcp-excel",
  "version": "0.1.0",
  "description": "让 AI 助手直接读写本地 Excel 文件的 MCP 服务（支持 xls / xlsx / csv）",
  "type": "module",
  "bin": { "mcp-excel": "dist/index.js" },
  "files": ["dist"],
  "engines": { "node": ">=20" },
  "scripts": {
    "build": "tsc",
    "start": "node dist/index.js",
    "prepublishOnly": "npm run build"
  },
  "license": "MIT",
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.0.0",
    "exceljs": "^4.4.0",
    "xlsx": "^0.18.5",
    "zod": "^3.25.0"
  },
  "devDependencies": {
    "@types/node": "^20.0.0",
    "typescript": "^5.5.0"
  }
}
```

**关键点**：
- `bin` 指向编译产物，因此 `dist/index.js` 第一行必须是 shebang（见 2.5，源码第一行写上 `#!/usr/bin/env node`，tsc 会原样保留）。npm 安装 bin 时会自动处理可执行权限，Windows 下生成 `.cmd` 包装，无需手工处理。
- `files: ["dist"]` 保证发布包只含编译产物。
- 依赖 `xlsx` / `exceljs` / `zod` 自带类型声明，无需额外 `@types/*`。

**出错处理**：本文件无运行时行为。

---

### 2.2 `tsconfig.json`

**负责什么**：把 `src/` 编译为 ESM 产物到 `dist/`。

**完整内容**：

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true
  },
  "include": ["src"]
}
```

**关键点**：`module: NodeNext` 配合 `"type": "module"`，强制源码相对导入带 `.js` 后缀（全局约定 6），编译期就能发现导入路径错误。

**出错处理**：无。

---

### 2.3 `.gitignore`

**完整内容**：

```
node_modules/
dist/
```

---

### 2.4 `README.md`

**负责什么**：npm 包说明页全文，是用户唯一会"看"的页面。详细布局、文案、配置示例见本文档 **第 4.1 节**，编码时照抄第 4 章即可。

**依赖**：无。**出错处理**：无。

---

### 2.5 `src/index.ts`

**负责什么**：进程入口。AI 助手按 MCP 配置启动本进程后，本文件建立 stdio 通道并挂上全部 15 个工具，然后保持运行等待消息。

**完整骨架**（此文件短，直接给全量）：

```ts
#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerAllTools } from "./tools/index.js";

async function main(): Promise<void> {
  const server = new McpServer({
    name: "mcp-excel",
    version: "0.1.0",
  });
  registerAllTools(server);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("mcp-excel 已启动，通过 stdio 等待 AI 助手连接");
}

main().catch((e: unknown) => {
  console.error("mcp-excel 启动失败：", e);
  process.exit(1);
});
```

**包含的函数**：

| 函数 | 输入 | 输出 | 说明 |
|------|------|------|------|
| `main()` | 无 | `Promise<void>` | 创建 server → 注册工具 → 连 stdio |

**依赖**：`@modelcontextprotocol/sdk`、`./tools/index.js`。

**出错处理**：
- `main()` 的任何 rejection（如 stdio 被占用、SDK 初始化失败）→ 打印到 stderr 并 `process.exit(1)`。AI 助手会看到进程退出并提示用户。
- 运行期单个工具的错误不经过本文件（由工具层的 `tool()` 包装器处理），本文件不加 `uncaughtException` 之外的额外逻辑。
- **严禁**在本文件（及全项目）使用 `console.log`（全局约定 3）。

---

### 2.6 `src/tools/index.ts`

**负责什么**：汇总注册。本身不含任何工具逻辑。

**包含的函数**：

| 函数 | 输入 | 输出 | 说明 |
|------|------|------|------|
| `registerAllTools(server: McpServer): void` | MCP server 实例 | 无 | 依次调用下面 6 个注册函数 |

**实现**（唯一逻辑）：

```ts
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerWorkbookTools } from "./workbook.js";
import { registerReadTools } from "./read.js";
import { registerWriteTools } from "./write.js";
import { registerFormatTools } from "./format.js";
import { registerDataTools } from "./data.js";
import { registerImageTools } from "./image.js";

export function registerAllTools(server: McpServer): void {
  registerWorkbookTools(server);
  registerReadTools(server);
  registerWriteTools(server);
  registerFormatTools(server);
  registerDataTools(server);
  registerImageTools(server);
}
```

**依赖**：`./workbook.js` `./read.js` `./write.js` `./format.js` `./data.js` `./image.js`。

**出错处理**：无（注册失败属于启动失败，由 2.5 的 catch 兜底）。

---

### 2.7 `src/tools/common.ts`

**负责什么**：所有工具文件共用的地基：zod 参数 schema、返回包络构造函数、错误包装器。

**导出的成员**：

**(1) zod schema（每个都带 `.describe()` 中文说明，这些说明就是给 AI 看的参数文档）：**

| 导出 | 定义 | describe 要点 |
|------|------|--------------|
| `filePathSchema` | `z.string().min(1)` | "Excel 文件的绝对路径，例如 `C:\\Users\\xx\\报表.xlsx` 或 `/Users/xx/报表.xlsx`。支持 .xlsx / .xls / .csv" |
| `sheetNameSchema` | `z.string().min(1).max(31).regex(/^[^\[\]\:\*\?\/\\]+$/)` | "工作表名称，1~31 个字符，不能包含 [ ] : * ? / \\" |
| `cellSchema` | `z.string().regex(/^[A-Za-z]+[0-9]+$/)` | "单元格地址，如 `B5`" |
| `rangeSchema` | `z.string().regex(/^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/)` | "矩形区域，如 `A1:D20`" |
| `valuesSchema` | `z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).min(1)` | "二维数组，每个内层数组是一行。`null` 表示清空该单元格" |

> sheetName 的正则与长度限制是 Excel 自身的规则，提前在参数层拦住，避免 ExcelJS 抛英文错。

**(2) 返回包络：**

```ts
import { toUserMessage } from "../engine/errors.js";

// 成功：{ success: true, ...data } 序列化为文本内容
export function ok(data: Record<string, unknown>): McpTextResult;

// 失败：任何异常 → 人话 → { success: false, error }，并置 isError
export function fail(e: unknown): McpTextResult;
```

其中 `McpTextResult` 为 `{ content: [{ type: "text"; text: string }]; isError?: boolean }`（类型可直接从 SDK 的 handler 返回值推断，用 `satisfies` 或显式定义均可）。序列化用 `JSON.stringify(obj, null, 2)`——带缩进，方便 AI 和用户阅读。

**(3) `tool()` 包装器（全局约定 4/5 的落点）：**

```ts
export function tool<A extends Record<string, unknown>>(
  handler: (args: A) => Promise<Record<string, unknown>> | Record<string, unknown>,
): (args: A) => Promise<McpTextResult>;
```

- 输入：业务 handler（返回"结果字段对象"，不含 `success`）。
- 输出：符合 SDK 要求的 handler——内部 `try { return ok(await handler(args)) } catch (e) { return fail(e) }`。
- 效果：每个业务 handler 只管抛 `ToolError` 或返回结果，永远不需要自己 try/catch。

**依赖**：`zod`、`../engine/errors.js`。

**出错处理**：`fail()` 内调用 `toUserMessage()`，保证任何异常（包括库抛的英文异常）都变成中文人话，绝不把堆栈暴露给 AI。

---

### 2.8 `src/tools/workbook.ts`

**负责什么**：注册 5 个"工作簿/工作表管理"工具。导出 `registerWorkbookTools(server: McpServer): void`。

**依赖**：`./common.js`、`../engine/reader.js`（`readWorkbook`）、`../engine/writer.js`（`createWorkbook`、`editWorkbook`）、`../engine/errors.js`。

**注册方式统一为**（其余工具文件同此模式，后面不再重复）：

```ts
server.registerTool("工具名", {
  title: "中文标题",
  description: "中文描述：什么时候用、注意什么",
  inputSchema: { /* zod 字段形状 */ },
}, tool(async (args) => { /* 业务逻辑，返回结果字段对象 */ }));
```

---

**(1) `create_workbook` —— 新建 Excel 文件**

- **description 要点**："从无到有新建一个 .xlsx 文件。可指定初始工作表名列表。文件已存在时会失败（防误覆盖），不会清空已有文件。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | 必须以 `.xlsx` 结尾，否则报 `INVALID_PARAMS`（"新建文件的扩展名必须是 .xlsx"） |
| `sheets` | `z.array(sheetNameSchema).min(1).optional()` | 否 | 初始工作表名；不传则创建一张 `"Sheet1"`；数组内重名报 `INVALID_PARAMS` |

- **实现步骤**：校验扩展名与重名 → `createWorkbook(filePath, sheets ?? ["Sheet1"])` → 返回 `{ sheetsCreated: n }`。
- **成功返回**：`{ success: true, sheetsCreated: 12 }`（`sheetsCreated` 为实际创建张数）。
- **错误**：`FILE_EXISTS`（目标已存在）、`DIR_NOT_FOUND`（目录不存在）、`INVALID_PARAMS`（扩展名/重名/表名非法）。

**(2) `list_sheets` —— 列出所有工作表**

- **description 要点**："查看一个 Excel 文件里有哪些工作表，各有多少行多少列。动手读/写之前可以先调它确认表名。"
- **输入**：仅 `filePath`。
- **实现步骤**：`readWorkbook(filePath)` → 把 `sheets` 映射为 `{ name, rowCount, colCount }`。
- **成功返回**：`{ success: true, sheets: [{ name, rowCount, colCount }] }`。
- **错误**：`FILE_NOT_FOUND`、`UNSUPPORTED_FORMAT`、文件损坏（`UNKNOWN`）。

**(3) `add_sheet` —— 新增一张工作表**

- **description 要点**："在已有文件中追加一张空工作表。"
- **输入**：`filePath`、`sheetName`（`sheetNameSchema`）。
- **实现步骤**：`editWorkbook(filePath, wb => { 若 wb.getWorksheet(sheetName) 存在 → 抛 ToolError("SHEET_EXISTS")；wb.addWorksheet(sheetName); })`。
- **成功返回**：`{ success: true }`。
- **错误**：`SHEET_EXISTS`、`READ_ONLY_FORMAT`（.xls）、`UNSUPPORTED_FORMAT`（.csv，文案"CSV 文件只有一张工作表，不支持新增"）。

**(4) `delete_sheet` —— 删除一张工作表**

- **description 要点**："删除指定工作表。不能删除最后一张。"
- **输入**：`filePath`、`sheetName`。
- **实现步骤**：`editWorkbook` 内：找不到 → `SHEET_NOT_FOUND`（附现有表名）；`wb.worksheets.length === 1` 且就是要删的这张 → `LAST_SHEET`；否则 `wb.removeWorksheet(ws.id)`。
- **成功返回**：`{ success: true }`。
- **错误**：`SHEET_NOT_FOUND`、`LAST_SHEET`、`READ_ONLY_FORMAT`、`UNSUPPORTED_FORMAT`（csv 同上报错）。

**(5) `rename_sheet` —— 重命名工作表**

- **description 要点**："修改工作表名称。"
- **输入**：`filePath`、`oldName`（`z.string().min(1)`）、`newName`（`sheetNameSchema`）。
- **实现步骤**：`editWorkbook` 内：`oldName` 找不到 → `SHEET_NOT_FOUND`；`newName` 已被占用（且不等于 oldName）→ `SHEET_EXISTS`；否则 `ws.name = newName`。
- **成功返回**：`{ success: true }`。
- **错误**：同上两类格式错误 + `SHEET_NOT_FOUND` / `SHEET_EXISTS`。

**csv 特例**：(3)(4)(5) 对 csv 一律由 `editWorkbook` 前的格式检查拦住（见 2.18 规格：工作表管理类操作传 `{ allowSheetOps: false }` 之类的标记不是必须的——统一规则是：**工具先按全局约定 8 的矩阵自己检查扩展名并抛错**，再调引擎。csv 判断用 `path.extname(filePath).toLowerCase() === ".csv"`，xls 判断同理由 `editWorkbook` 内部拦截。）

---

### 2.9 `src/tools/read.ts`

**负责什么**：注册 `read_range`。导出 `registerReadTools(server)`。

**依赖**：`./common.js`、`../engine/reader.js`、`../engine/address.js`。

**`read_range` —— 读取区域内容**

- **description 要点**："读取一张工作表的内容，返回二维数组。不传 range 则读整张表；表很大时建议先传 range 读一部分（如前 50 行）了解结构。空单元格为 null，日期为 ISO 字符串。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | |
| `sheetName` | `z.string().min(1).optional()` | 否 | 默认第一张工作表 |
| `range` | `rangeSchema.optional()` | 否 | 如 `"A1:D20"`；不传 = 整张表 |

- **实现步骤**：
  1. `readWorkbook(filePath)` 得到 `UnifiedWorkbook`。
  2. 定位 sheet：`sheetName` 未传取 `sheets[0]`；传了但找不到 → `SHEET_NOT_FOUND`（错误文案附现有表名列表）。
  3. 未传 `range` → 直接返回 `sheet.values`。
  4. 传了 `range` → `parseRange()` 解析，遍历矩形内每个坐标：落在 `values` 范围内取实际值，超出补 `null`。即返回数组行数恰为请求行数、列数恰为请求列数。
- **成功返回**：`{ success: true, values: CellValue[][] }`。
- **错误**：`FILE_NOT_FOUND`、`UNSUPPORTED_FORMAT`、`SHEET_NOT_FOUND`、`INVALID_RANGE`。
- **空表**：成功返回 `values: []`（不是错误，见 4.2 "没数据"状态）。

---

### 2.10 `src/tools/write.ts`

**负责什么**：注册 `write_range`、`set_formula`。导出 `registerWriteTools(server)`。

**依赖**：`./common.js`、`../engine/writer.js`、`../engine/address.js`。

---

**(1) `write_range` —— 写入一块数据（单格写入 = 传 1×1 数组，不另设工具）**

- **description 要点**："从 startCell 开始写入一个二维数组（先行后列）。一个格子就传 `[[值]]`；一大片就传多行多列。值为 null 会清空对应单元格。字符串不会被套公式——要写公式请用 set_formula。目标工作表必须已存在（没有就先用 add_sheet）。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | |
| `sheetName` | `z.string().min(1)` | ✅ | csv 文件忽略此参数，直接写唯一工作表 |
| `startCell` | `cellSchema` | ✅ | 区域左上角，如 `"B5"` |
| `values` | `valuesSchema` | ✅ | 各行长度可以不一致（按实际长度写） |

- **实现步骤**：`parseCell(startCell)` → `editWorkbook(filePath, wb => { const ws = getSheetOrThrow(wb, sheetName); 双重循环：ws.getCell(row + i, col + j).value = v; })`（`v === null` 赋 `null` 即清空）。统计处理过的单元格总数（含 null）。
- **成功返回**：`{ success: true, cellsWritten: number }`。
- **错误**：`SHEET_NOT_FOUND`、`INVALID_CELL`、`READ_ONLY_FORMAT`（.xls）、`FILE_NOT_FOUND`、`FILE_BUSY`。

**(2) `set_formula` —— 在单元格写入公式**

- **description 要点**："在指定单元格写入 Excel 公式，如 `SUM(A1:A10)`，可带或不带开头的 `=`。注意：公式要等用户在 Excel/WPS 里打开文件后才会计算出结果，本工具读该格会得到 null。"
- **输入**：`filePath`、`sheetName`、`cell`（`cellSchema`）、`formula`（`z.string().min(1)`，describe："公式文本，如 `SUM(A1:A10)` 或 `=AVERAGE(B2:B30)`"）。
- **实现步骤**：去掉 `formula` 开头的 `=` → `editWorkbook` 内 `ws.getCell(row, col).value = { formula }`。
- **成功返回**：`{ success: true }`。
- **错误**：同 write_range + `INVALID_PARAMS`（去 `=` 后为空串）。

---

### 2.11 `src/tools/format.ts`

**负责什么**：注册 `format_cells`、`merge_cells`、`set_dimensions`。导出 `registerFormatTools(server)`。

**依赖**：`./common.js`、`../engine/writer.js`、`../engine/address.js`。

---

**(1) `format_cells` —— 设置样式**

- **description 要点**："设置一个区域的字体、填充色、边框、对齐、数字格式，让表格好看。只传想改的项，没传的不动。csv 文件不支持样式。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | |
| `sheetName` | `z.string().min(1)` | ✅ | |
| `range` | `rangeSchema` | ✅ | |
| `style` | 见下 | ✅ | 至少含一个子字段（zod `.refine` 校验，否则报 `INVALID_PARAMS`） |

`style` 的 zod 结构（全部可选）：

```ts
const styleSchema = z.object({
  font: z.object({
    name: z.string().optional(),          // 字体名，如 "微软雅黑"
    size: z.number().positive().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    color: z.string().optional(),         // 6 位 hex，如 "FF0000"，可带 # 前缀
  }).optional(),
  fill: z.object({
    color: z.string().optional(),         // 纯色填充，6 位 hex
  }).optional(),
  border: z.object({
    style: z.enum(["thin", "medium", "thick"]).optional(),
    color: z.string().optional(),
  }).optional(),                          // 四边统一设置
  alignment: z.object({
    horizontal: z.enum(["left", "center", "right"]).optional(),
    vertical: z.enum(["top", "middle", "bottom"]).optional(),
    wrapText: z.boolean().optional(),
  }).optional(),
  numberFormat: z.string().optional(),    // 如 "0.00"、"yyyy-mm-dd"、"#,##0"
}).refine(o => Object.values(o).some(v => v !== undefined), { message: "style 至少要包含一项设置" });
```

- **实现步骤**：`parseRange()` → `editWorkbook` 内遍历区域每个 `cell`：
  - `font` → `cell.font = { ...font, color: font.color ? { argb: toArgb(font.color) } : undefined }`
  - `fill` → `cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: toArgb(color) } }`
  - `border` → `cell.border = { top/left/bottom/right: { style, color? } }`（四边同一对象）
  - `alignment` → `cell.alignment = alignment`
  - `numberFormat` → `cell.numFmt = numberFormat`
  - `toArgb(hex)`：去掉 `#`，转大写，前面补 `"FF"` alpha。
- **成功返回**：`{ success: true }`。
- **错误**：`SHEET_NOT_FOUND`、`INVALID_RANGE`、`INVALID_PARAMS`（style 为空）、`READ_ONLY_FORMAT`、`UNSUPPORTED_FORMAT`（csv，"CSV 不支持样式"）。

**(2) `merge_cells` —— 合并单元格**

- **description 要点**："把一个矩形区域合并成一个单元格，内容保留左上角的值。常用于标题行。"
- **输入**：`filePath`、`sheetName`、`range`（`rangeSchema`，必须跨越多于 1 个单元格，否则报 `INVALID_PARAMS`）。
- **实现步骤**：`editWorkbook` 内 `ws.mergeCells(startRow, startCol, endRow, endCol)`（ExcelJS 此 API 为 1 起始，直接用解析结果）。
- **成功返回**：`{ success: true }`。
- **错误**：同上；与已有合并区域重叠时 ExcelJS 可能抛错 → 由 `fail()` 兜底成 `UNKNOWN` 人话。

**(3) `set_dimensions` —— 设置列宽行高**

- **description 要点**："设置工作表的列宽和/或行高，让表格排版合适。列宽单位是字符数（Excel 默认约 8.43），行高单位是磅（默认约 15）。columns 和 rows 至少传一个。csv 不支持。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | |
| `sheetName` | `z.string().min(1)` | ✅ | |
| `columns` | `z.array(z.object({ column: z.string().regex(/^[A-Za-z]+$/), width: z.number().positive() })).min(1).optional()` | 否 | 每项 = 列字母（如 `"B"`，不区分大小写）+ 列宽 |
| `rows` | `z.array(z.object({ row: z.number().int().positive(), height: z.number().positive() })).min(1).optional()` | 否 | 每项 = 行号（1 起始）+ 行高 |

- **实现步骤**：`columns` / `rows` 都未传 → `INVALID_PARAMS`（"columns 和 rows 至少要传一个"）→ `editWorkbook` 内定位工作表（找不到 → `SHEET_NOT_FOUND`）→ 每列 `ws.getColumn(letterToCol(column)).width = width`；每行 `ws.getRow(row).height = height`。
- **成功返回**：`{ success: true }`。
- **错误**：`SHEET_NOT_FOUND`、`INVALID_PARAMS`、`READ_ONLY_FORMAT`（.xls）、`UNSUPPORTED_FORMAT`（csv，"CSV 文件不支持设置列宽行高。"）。

> 说明：列宽/行高是列级/行级属性而非单元格样式，因此独立设工具而不并入 `format_cells`；`format_cells` 的 "range + style" 语义保持不变。

---

### 2.12 `src/tools/data.ts`

**负责什么**：注册 `sort_range`、`filter_range`、`dedupe_range`。导出 `registerDataTools(server)`。

**依赖**：`./common.js`、`../engine/reader.js`（filter）、`../engine/writer.js`（sort/dedupe 的 `editWorkbook` 与 `normalizeCellValue`）、`../engine/address.js`。

**三个工具共用的列标识规则**（实现集中在 `resolveColumn()`，见 2.16）：
- `keyColumn` / `column` / `keyColumns` 接受两种写法：**列字母**（如 `"C"`，指工作表绝对列，不区分大小写）或 **表头名**（如 `"销售额"`，仅当 `hasHeader: true` 时按第一行精确匹配）。
- 解析失败抛 `INVALID_PARAMS`，文案列明可用列（如 `列 "销量" 不存在。表头包含：姓名, 销售额, 日期`）。
- `hasHeader` 统一为 `z.boolean().optional()`，**默认 false**；为 true 时第一行视为表头，不参与排序/去重，filter 时置于结果第一行原样返回。

**值比较规则**（`compareValues(a, b)`，实现于本文件，供 sort 用）：`null` 永远排最后；两个都是 number 比数值；其余情况 `String(a).localeCompare(String(b), "zh-Hans-CN")`；`desc` 时取反（null 仍然最后）。

---

**(1) `sort_range` —— 排序**

- **description 要点**："对一个区域按某列排序（会修改文件）。区域外的内容不动。有表头时传 hasHeader: true 并用表头名指定列更稳妥。"
- **输入**：`filePath`、`sheetName`、`range`、`keyColumn`（`z.string().min(1)`）、`order`（`z.enum(["asc", "desc"])`）、`hasHeader?`。
- **实现步骤**（`editWorkbook` 内）：
  1. 把区域每格的**原始 `cell.value`**（不归一化，保住日期/数字类型）读成 `raw[][]`；另算一份 `normalizeCellValue` 后的值用于比较。
  2. 解析 `keyColumn` → 区域内列下标。
  3. 数据行（hasHeader 时跳过第 0 行）按 `compareValues` 稳定排序（`Array.prototype.sort` 在 V8 中是稳定排序）。
  4. 按新顺序把 `raw` 写回原区域。
- **成功返回**：`{ success: true }`。
- **错误**：`SHEET_NOT_FOUND`、`INVALID_RANGE`、`INVALID_PARAMS`（列不存在）、`READ_ONLY_FORMAT`、`UNSUPPORTED_FORMAT`（csv 支持排序——它是值操作，见全局约定 8；样式/表操作才拦 csv）。

**(2) `filter_range` —— 按条件筛选（只读，不改文件）**

- **description 要点**："按条件筛选区域中的行并返回结果，不修改文件。多个条件为'并且'关系。"
- **输入**：`filePath`、`sheetName`、`range`、`conditions`（`z.array(conditionSchema).min(1)`）、`hasHeader?`。

```ts
const conditionSchema = z.object({
  column: z.string().min(1),   // 列字母或表头名
  op: z.enum(["=", "!=", ">", ">=", "<", "<=", "contains"]),
  value: z.union([z.string(), z.number(), z.boolean()]),
});
```

- **实现步骤**：`readWorkbook()` → 定位 sheet → `parseRange` 裁剪出块 → 表头行（若有）用于列解析且永远排在结果第一行 → 逐行检查全部条件：
  - `=` / `!=`：宽松相等——先都按原类型比，不等再比 `String()` 形式；
  - `>` / `>=` / `<` / `<=`：两侧都能转为有限 number 时比数值，否则比 `localeCompare` 字符串序；
  - `contains`：`String(cell).includes(String(value))`，大小写敏感。
- **成功返回**：`{ success: true, values: CellValue[][] }`（0 行匹配且 `hasHeader: true` 时返回仅含表头的数组；无表头 0 匹配返回 `[]`——这是"没数据"，不是错误）。
- **错误**：`SHEET_NOT_FOUND`、`INVALID_RANGE`、`INVALID_PARAMS`、`.xls` **允许**（只读操作）。

**(3) `dedupe_range` —— 去重**

- **description 要点**："删除区域内 keyColumns 完全相同的重复行，只保留第一次出现的行（会修改文件）。被删的行变为空白行，区域外的内容不会被移动。"
- **输入**：`filePath`、`sheetName`、`range`、`keyColumns`（`z.array(z.string().min(1)).min(1)`）、`hasHeader?`。
- **实现步骤**（`editWorkbook` 内）：
  1. 同 sort 读出 `raw[][]` 与归一化值。
  2. 逐数据行算 key：各 key 列归一化值 `JSON.stringify` 后拼接；`Set` 判重，保留首现。
  3. 把保留行紧凑写回区域前部；区域尾部多出来的行整行写 `null`（不 `spliceRows`，避免移动区域外内容——这是有意的取舍）。
- **成功返回**：`{ success: true, removedCount: number }`。
- **错误**：同 sort_range。

---

### 2.13 `src/tools/image.ts`

**负责什么**：注册 `insert_image`。导出 `registerImageTools(server)`。

**依赖**：`./common.js`、`../engine/writer.js`、`../engine/address.js`、`node:fs`、`node:path`。

**`insert_image` —— 把图片插入工作表**

- **description 要点**："把一张本地图片（png / jpg / gif）插入工作表，图片左上角锚定在 anchorCell。width/height 为像素，不传给默认 300×200——比例不对时请显式传这两个值。csv 不支持插图。"
- **输入**：

| 字段 | schema | 必填 | 说明 |
|------|--------|:---:|------|
| `filePath` | `filePathSchema` | ✅ | |
| `sheetName` | `z.string().min(1)` | ✅ | |
| `imagePath` | `z.string().min(1)` | ✅ | 图片绝对路径 |
| `anchorCell` | `cellSchema` | ✅ | 图片左上角放在哪个格子 |
| `width` | `z.number().positive().optional()` | 否 | 像素，默认 300 |
| `height` | `z.number().positive().optional()` | 否 | 像素，默认 200 |

> 默认尺寸定为 **300×200**（不读图片原始尺寸，避免引入额外依赖；description 里已提示 AI 显式传尺寸）。

- **实现步骤**：
  1. `fs.existsSync(imagePath)` 失败 → `IMAGE_NOT_FOUND`。
  2. 扩展名映射：`.png`→`"png"`、`.jpg`/`.jpeg`→`"jpeg"`、`.gif`→`"gif"`，其他 → `UNSUPPORTED_IMAGE`。
  3. `editWorkbook` 内：
     ```ts
     const imageId = wb.addImage({ filename: imagePath, extension });
     ws.addImage(imageId, {
       tl: { col: col - 1, row: row - 1 },   // ExcelJS 的 tl 是 0 起始！
       ext: { width: width ?? 300, height: height ?? 200 },
     });
     ```
- **成功返回**：`{ success: true }`。
- **错误**：`IMAGE_NOT_FOUND`、`UNSUPPORTED_IMAGE`、`SHEET_NOT_FOUND`、`INVALID_CELL`、`READ_ONLY_FORMAT`、`UNSUPPORTED_FORMAT`（csv）。

---

### 2.14 `src/engine/types.ts`

**负责什么**：定义引擎层与工具层之间流转的统一内存结构。纯类型文件，无任何运行时代码、无依赖。

**导出的类型**：

```ts
/** 单元格值的统一形态：JSON 可序列化的四种 */
export type CellValue = string | number | boolean | null;

/** 文件格式（由扩展名判定） */
export type BookFormat = "xlsx" | "xls" | "csv";

export interface UnifiedSheet {
  name: string;
  rowCount: number;              // values.length（已裁掉尾部全空行）
  colCount: number;              // 所有行中的最大列数（空表为 0）
  values: CellValue[][];         // 每行长度已补齐到 colCount
}

export interface UnifiedWorkbook {
  format: BookFormat;
  sheets: UnifiedSheet[];
}
```

**值的归一化规则**（`reader.ts` 按此产出、`filter` 等按此消费）：
- `Date` → ISO 字符串（`toISOString()`，如 `"2026-01-05T00:00:00.000Z"`）。
- `string` / `number` / `boolean` 原样保留。
- 公式格：取缓存计算结果；无缓存结果 → `null`。
- 其他一切（undefined、富文本、错误值等）→ `null`。

**出错处理**：无。

---

### 2.15 `src/engine/errors.ts`

**负责什么**：全项目唯一的错误类型与"异常 → 人话"转换。

**导出的成员**：

**(1) 错误码与 `ToolError`：**

```ts
export type ErrorCode =
  | "FILE_NOT_FOUND" | "DIR_NOT_FOUND" | "FILE_EXISTS" | "FILE_BUSY"
  | "UNSUPPORTED_FORMAT" | "READ_ONLY_FORMAT"
  | "SHEET_NOT_FOUND" | "SHEET_EXISTS" | "LAST_SHEET"
  | "INVALID_CELL" | "INVALID_RANGE" | "INVALID_PARAMS"
  | "IMAGE_NOT_FOUND" | "UNSUPPORTED_IMAGE"
  | "UNKNOWN";

export class ToolError extends Error {
  readonly code: ErrorCode;
  constructor(code: ErrorCode, message: string) { super(message); this.code = code; }
}
```

`message` 直接就是给最终用户看的中文文案（含具体路径、表名等上下文），抛错处按第 3.4 节的文案模板拼装。

**(2) `toUserMessage(e: unknown): string`** —— `common.ts` 的 `fail()` 调它：

| 输入 | 返回 |
|------|------|
| `ToolError` | `e.message` 原样返回 |
| `z.ZodError` | `"参数错误：" + 第一条 issue 的 message` |
| `Error` 且 `code === "ENOENT"` | `"文件或目录不存在：" + 原始 path 信息` |
| `Error` 且 `code` 为 `EPERM` / `EACCES` / `EBUSY` | `"文件正被其他程序（如 Excel）占用，请关闭后重试。"` |
| 其他 `Error` | `"操作失败：" + e.message` |
| 非 Error | `"操作失败：未知错误"` |

**依赖**：无项目内依赖（`zod` 仅作类型引用，可用 `instanceof` 判断）。

**出错处理**：本文件自身不抛错。

---

### 2.16 `src/engine/address.ts`

**负责什么**：所有"地址/列"解析的唯一实现（全局约定 9）。纯函数，只依赖 `./errors.js`。

**导出的成员**：

```ts
export interface CellAddr { row: number; col: number }          // 均 1 起始
export interface RangeAddr { start: CellAddr; end: CellAddr }   // start 恒在左上、end 恒在右下

export function parseCell(input: string): CellAddr;
// "B5" → { row: 5, col: 2 }。正则不匹配或行号为 0 → ToolError("INVALID_CELL")

export function parseRange(input: string): RangeAddr;
// "A1:D20" → 两端各 parseCell；若顺序颠倒（如 "D20:A1"）自动交换为左上→右下，不报错。
// 格式不合法 → ToolError("INVALID_RANGE")

export function toCellName(addr: CellAddr): string;   // { row: 5, col: 2 } → "B5"
export function colToLetter(col: number): string;     // 1 → "A"，28 → "AB"
export function letterToCol(letters: string): number; // "ab" → 28（不区分大小写）

export function resolveColumn(
  token: string,
  range: RangeAddr,
  headerRow?: CellValue[],   // hasHeader 时传区域第一行
): number;                     // 返回工作表绝对列号（1 起始）
```

**`resolveColumn` 逻辑**（sort/filter/dedupe 共用）：
1. `token` 全是字母 → `letterToCol(token)`；结果必须落在 `[range.start.col, range.end.col]` 内，否则抛 `INVALID_PARAMS`（"列 X 不在区域 A1:D20 内"）。
2. 否则视为表头名：要求 `headerRow` 已提供（否则抛 `INVALID_PARAMS`："使用表头名指定列时，请设置 hasHeader: true"）；在 `headerRow` 中按 `String(v) === token` 精确匹配，命中返回 `range.start.col + 下标`；未命中抛 `INVALID_PARAMS`（文案附全部表头）。

**出错处理**：全部通过抛 `ToolError` 表达，不返回 undefined。

---

### 2.17 `src/engine/reader.ts`

**负责什么**：读取适配——用 SheetJS 把任意支持的文件读成 `UnifiedWorkbook`。全项目**唯一** import `xlsx` 的地方。

**导出的成员**：

```ts
export function readWorkbook(filePath: string): UnifiedWorkbook;  // 同步（SheetJS readFile 本身是同步 IO）
```

**实现步骤**：
1. `fs.existsSync(filePath)` 失败 → `ToolError("FILE_NOT_FOUND")`（文案含路径，提示需要绝对路径）。
2. `path.extname(filePath).toLowerCase()` 判定格式：`.xlsx`/`.xlsm` → `"xlsx"`；`.xls` → `"xls"`；`.csv` → `"csv"`；其余 → `UNSUPPORTED_FORMAT`（文案："仅支持 .xlsx / .xls / .csv"）。
3. `XLSX.readFile(filePath, { cellDates: true })`。抛异常 → 包装为 `ToolError("UNKNOWN", "无法解析文件：…，文件可能已损坏或不是有效的表格文件。")`。
4. 对每个 `wb.SheetNames`：`XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })` 得到 `unknown[][]`，然后归一化：
   - 逐值转换按 2.14 的规则（`v instanceof Date → v.toISOString()`；`string/number/boolean` 保留；其余 → `null`）。
   - `colCount` = 各行最大长度；每行 `push(null)` 补齐到 `colCount`。
   - 裁掉**尾部**全空行（整行皆 `null` 或空数组）；中间的空白行保留。
   - `rowCount` = 裁剪后行数。空表得到 `{ rowCount: 0, colCount: 0, values: [] }`。
5. 返回 `{ format, sheets }`。

**csv 说明**：SheetJS 读 csv 会得到单工作表（名通常为 `"Sheet1"`），走同一条管线，天然满足"读 csv"。

**依赖**：`xlsx`、`node:fs`、`node:path`、`./types.js`、`./errors.js`。

**出错处理**：见步骤 1/2/3，全部 `ToolError`。

---

### 2.18 `src/engine/writer.ts`

**负责什么**：编辑适配——用 ExcelJS 新建/修改/保存文件，含原子写。全项目**唯一** import `exceljs` 的地方。

**导出的成员**：

```ts
import type ExcelJS from "exceljs";

/** 新建一个 xlsx 工作簿并保存。文件已存在 → ToolError("FILE_EXISTS") */
export async function createWorkbook(filePath: string, sheetNames: string[]): Promise<void>;

/**
 * 打开 → 执行 mutator → 原子保存。
 * - 文件不存在 → FILE_NOT_FOUND
 * - .xls → READ_ONLY_FORMAT（ExcelJS 不支持 xls）
 * - .xlsx/.xlsm → workbook.xlsx.readFile；.csv → workbook.csv.readFile（单表）
 * mutator 内抛的 ToolError 原样向上抛；抛其他异常 → UNKNOWN。
 */
export async function editWorkbook(
  filePath: string,
  mutator: (wb: ExcelJS.Workbook) => void | Promise<void>,
): Promise<void>;

/** ExcelJS 原始 cell.value → 统一 CellValue（sort/dedupe 比较用，见 2.12） */
export function normalizeCellValue(v: ExcelJS.CellValue): CellValue;
```

**`createWorkbook` 实现步骤**：
1. `fs.existsSync` → `FILE_EXISTS`（"文件已存在，如需修改请用写入类工具，如需重建请先删除"）。
2. `new ExcelJS.Workbook()`，逐个 `addWorksheet(name)`。
3. `atomicSave(wb, filePath, "xlsx")`。

**`editWorkbook` 实现步骤**：
1. 不存在 → `FILE_NOT_FOUND`。
2. 扩展名：`.xls` → `READ_ONLY_FORMAT`（"老版 .xls 不支持直接修改，请先在 Excel/WPS 中另存为 .xlsx"）；`.xlsx`/`.xlsm` → `await wb.xlsx.readFile(filePath)`；`.csv` → `await wb.csv.readFile(filePath)`；其他 → `UNSUPPORTED_FORMAT`。读取异常 → `UNKNOWN`（"无法解析文件…"）。
3. `await mutator(wb)`。
4. `atomicSave(wb, filePath, format)`（csv 时 ExcelJS 只写第一张工作表，可接受——csv 本就单表）。

**`atomicSave(wb, filePath, format)`（内部函数，不导出）**：
1. `tmpPath = `${filePath}.tmp-${process.pid}``。
2. `await wb.xlsx.writeFile(tmpPath)` 或 `await wb.csv.writeFile(tmpPath)`；`ENOENT` → `DIR_NOT_FOUND`（"保存失败：目录不存在"）。
3. `fs.renameSync(tmpPath, filePath)`（Windows 下覆盖已存在文件是安全的）；`EPERM`/`EACCES`/`EBUSY` → `FILE_BUSY`（"文件正被 Excel 等程序占用，请关闭后重试"）。
4. `finally`：若 `tmpPath` 仍存在则尽力删除（忽略错误）。

**`normalizeCellValue(v)` 映射表**：

| ExcelJS 原始值 | 输出 |
|---------------|------|
| `null` / `undefined` | `null` |
| `string` / `number` / `boolean` | 原样 |
| `Date` | `toISOString()` |
| `{ richText: [...] }` | 各段 `text` 拼接 |
| `{ text, hyperlink }` | `text` |
| `{ formula, result }` / `{ sharedFormula, result }` | `result` 归一化后返回，无 `result` → `null` |
| `{ error }` | `null` |
| 其他对象 | `null` |

**依赖**：`exceljs`、`node:fs`、`node:path`、`./types.js`、`./errors.js`。

**出错处理**：见上；本文件不吞任何 `ToolError`——mutator 抛的原样透传给工具层的 `tool()` 包装器。

---

## 3. 数据格式定义

### 3.1 内存结构（`UnifiedWorkbook`）

工具层与引擎层之间的唯一数据形态（定义见 2.14）。具体示例——《2026年销售记录.xlsx》的"一月"表被 `readWorkbook()` 读入后：

```json
{
  "format": "xlsx",
  "sheets": [
    {
      "name": "一月",
      "rowCount": 4,
      "colCount": 4,
      "values": [
        ["姓名", "销售额", "日期", "备注"],
        ["张三", 1200, "2026-01-05T00:00:00.000Z", null],
        ["李四", 980.5, "2026-01-06T00:00:00.000Z", "已回款"],
        [null, null, null, null]
      ]
    },
    {
      "name": "二月",
      "rowCount": 0,
      "colCount": 0,
      "values": []
    }
  ]
}
```

规则复述（细节见 2.14/2.17）：
- `CellValue = string | number | boolean | null`，永远是 JSON 可序列化的。
- 日期一律是 ISO 字符串（上面的 `"2026-01-05T00:00:00.000Z"`），不是特殊类型。
- 每行长度补齐到 `colCount`；行中间的空白是 `null`；文件尾部全空行被裁掉（上面第 4 行是文件中间的空白行，保留）。
- 空工作表是合法的：`values: []`。

### 3.2 工具返回包络（MCP 消息层）

每个工具的 handler 返回值结构（SDK 约定）：

```json
{
  "content": [
    { "type": "text", "text": "{ \"success\": true, ... } 的 JSON 字符串（2 空格缩进）" }
  ]
}
```

- 成功：`text` 为 `{ "success": true, ...结果字段 }`，不设 `isError`。
- 失败：`text` 为 `{ "success": false, "error": "人话错误" }`，并设 `"isError": true`。
- handler 永远 resolve（全局约定 4）：协议层不会看到 exception，AI 总能拿到一段可解析的 JSON。

### 3.3 十五个工具的输入 / 输出 JSON 示例

> 输入 = AI 通过 MCP 传来的参数对象；输出 = `text` 字段里序列化的 JSON（此处已格式化展示）。
> 各工具共有的失败（`FILE_NOT_FOUND`、`SHEET_NOT_FOUND` 等）只在首次出现处给示例，其余引用 3.4 文案表。

**(1) `create_workbook`**

输入：
```json
{ "filePath": "C:\\Users\\li\\考勤表.xlsx", "sheets": ["一月", "二月", "三月"] }
```
成功输出：
```json
{ "success": true, "sheetsCreated": 3 }
```
失败输出（文件已存在）：
```json
{ "success": false, "error": "文件已存在：C:\\Users\\li\\考勤表.xlsx。如需修改请直接使用写入类工具；如需重建请先删除该文件。" }
```

**(2) `list_sheets`**

输入：
```json
{ "filePath": "C:\\Users\\wang\\2026年销售记录.xlsx" }
```
成功输出：
```json
{ "success": true, "sheets": [ { "name": "一月", "rowCount": 20, "colCount": 4 }, { "name": "二月", "rowCount": 0, "colCount": 0 } ] }
```

**(3) `read_range`**

输入（指定区域）：
```json
{ "filePath": "C:\\Users\\wang\\2026年销售记录.xlsx", "sheetName": "一月", "range": "A1:B3" }
```
成功输出（注意第 3 行超出实际数据，按请求大小补 null）：
```json
{ "success": true, "values": [["姓名", "销售额"], ["张三", 1200], [null, null]] }
```
输入（整表）：`{ "filePath": "...", "sheetName": "一月" }` → `values` 为整张表。
失败输出（表名打错）：
```json
{ "success": false, "error": "工作表 \"一月份\" 不存在。现有工作表：一月, 二月。" }
```

**(4) `write_range`**

输入（改单格：1×1 数组）：
```json
{ "filePath": "C:\\Users\\zhang\\费用表.xlsx", "sheetName": "费用表", "startCell": "B5", "values": [[3500]] }
```
成功输出：
```json
{ "success": true, "cellsWritten": 1 }
```
输入（批量写多行）：
```json
{ "filePath": "...", "sheetName": "一月", "startCell": "A1", "values": [["姓名", 1, 2, 3], ["张三", "出勤", "请假", null]] }
```
成功输出：
```json
{ "success": true, "cellsWritten": 8 }
```
失败输出（.xls 写入）：
```json
{ "success": false, "error": "老版 .xls 文件不支持直接修改。请先在 Excel/WPS 中另存为 .xlsx 后再操作。" }
```

**(5) `add_sheet`**

输入：`{ "filePath": "...", "sheetName": "四月" }`
成功输出：`{ "success": true }`
失败输出（重名）：
```json
{ "success": false, "error": "工作表 \"四月\" 已存在。" }
```

**(6) `delete_sheet`**

输入：`{ "filePath": "...", "sheetName": "草稿" }`
成功输出：`{ "success": true }`
失败输出（最后一张）：
```json
{ "success": false, "error": "不能删除最后一张工作表。" }
```

**(7) `rename_sheet`**

输入：`{ "filePath": "...", "oldName": "Sheet1", "newName": "一季度" }`
成功输出：`{ "success": true }`

**(8) `insert_image`**

输入：
```json
{ "filePath": "C:\\Users\\zhao\\商品清单.xlsx", "sheetName": "清单", "imagePath": "C:\\Users\\zhao\\product.png", "anchorCell": "D2", "width": 120, "height": 120 }
```
成功输出：`{ "success": true }`
失败输出（图片不存在）：
```json
{ "success": false, "error": "图片文件不存在：C:\\Users\\zhao\\product.png。请检查路径。" }
```

**(9) `set_formula`**

输入：
```json
{ "filePath": "...", "sheetName": "费用表", "cell": "B10", "formula": "SUM(B2:B9)" }
```
成功输出：`{ "success": true }`

**(10) `format_cells`**

输入（标题行美化：加粗、白字、深蓝底、居中）：
```json
{
  "filePath": "...", "sheetName": "一月", "range": "A1:D1",
  "style": {
    "font": { "bold": true, "color": "FFFFFF", "size": 12 },
    "fill": { "color": "305496" },
    "alignment": { "horizontal": "center", "vertical": "middle" },
    "border": { "style": "thin", "color": "999999" }
  }
}
```
成功输出：`{ "success": true }`
失败输出（csv）：
```json
{ "success": false, "error": "CSV 文件不支持样式设置。" }
```

**(11) `merge_cells`**

输入：`{ "filePath": "...", "sheetName": "一月", "range": "A1:D1" }`
成功输出：`{ "success": true }`

**(12) `sort_range`**

输入（按"销售额"列降序，第一行是表头）：
```json
{ "filePath": "...", "sheetName": "一月", "range": "A1:D20", "keyColumn": "销售额", "order": "desc", "hasHeader": true }
```
等价输入（用列字母）：`"keyColumn": "B"`。
成功输出：`{ "success": true }`
失败输出（列名打错）：
```json
{ "success": false, "error": "列 \"销量\" 不存在。表头包含：姓名, 销售额, 日期, 备注。" }
```

**(13) `filter_range`**

输入（找出销售额大于 1000 的行）：
```json
{ "filePath": "...", "sheetName": "一月", "range": "A1:D20", "hasHeader": true,
  "conditions": [ { "column": "销售额", "op": ">", "value": 1000 } ] }
```
成功输出（含表头行）：
```json
{ "success": true, "values": [["姓名", "销售额", "日期", "备注"], ["张三", 1200, "2026-01-05T00:00:00.000Z", null]] }
```
没数据输出（无匹配行，表头仍返回）：
```json
{ "success": true, "values": [["姓名", "销售额", "日期", "备注"]] }
```

**(14) `dedupe_range`**

输入（按姓名+日期去重）：
```json
{ "filePath": "...", "sheetName": "一月", "range": "A1:D20", "keyColumns": ["A", "C"], "hasHeader": true }
```
成功输出：
```json
{ "success": true, "removedCount": 3 }
```

**(15) `set_dimensions`**

输入（B 列加宽到 16、C 列到 20，第 1 行行高 24）：
```json
{ "filePath": "...", "sheetName": "一月", "columns": [ { "column": "B", "width": 16 }, { "column": "C", "width": 20 } ], "rows": [ { "row": 1, "height": 24 } ] }
```
成功输出：`{ "success": true }`
失败输出（csv）：
```json
{ "success": false, "error": "CSV 文件不支持设置列宽行高。" }
```

### 3.4 错误码与文案模板表

`ToolError` 的 `message` 按下表拼装（`{}` 为占位符）；`errors.ts` 的 `toUserMessage()` 负责兜住所有非 `ToolError` 异常（映射规则见 2.15）。

| 错误码 | 触发场景 | 文案模板 |
|--------|---------|---------|
| `FILE_NOT_FOUND` | 读/改不存在的文件 | `文件不存在：{filePath}。请检查路径是否正确（需要绝对路径）。` |
| `DIR_NOT_FOUND` | 保存时目录不存在 | `保存失败：目录不存在：{目录路径}。` |
| `FILE_EXISTS` | create_workbook 目标已存在 | `文件已存在：{filePath}。如需修改请直接使用写入类工具；如需重建请先删除该文件。` |
| `FILE_BUSY` | 保存时文件被占用 | `无法保存：{filePath} 正被其他程序（如 Excel）占用，请关闭后重试。` |
| `UNSUPPORTED_FORMAT` | 扩展名不支持 / csv 做表操作、样式、插图 | `不支持的文件格式：{ext}。仅支持 .xlsx / .xls / .csv。` / 按场景用 "CSV 文件只有一张工作表，不支持…" "CSV 文件不支持样式设置。" 等具体文案 |
| `READ_ONLY_FORMAT` | 对 .xls 执行写操作 | `老版 .xls 文件不支持直接修改。请先在 Excel/WPS 中另存为 .xlsx 后再操作。` |
| `SHEET_NOT_FOUND` | 工作表名不存在 | `工作表 "{name}" 不存在。现有工作表：{逗号连接的表名}。` |
| `SHEET_EXISTS` | 新增/重命名为已有表名 | `工作表 "{name}" 已存在。` |
| `LAST_SHEET` | 删除最后一张工作表 | `不能删除最后一张工作表。` |
| `INVALID_CELL` | 单元格地址不合法 | `无效的单元格地址："{input}"。正确示例："B5"。` |
| `INVALID_RANGE` | 区域不合法 | `无效的区域："{input}"。正确示例："A1:D20"。` |
| `INVALID_PARAMS` | 其他参数问题（列不存在、style 为空、扩展名须为 .xlsx 等） | 按场景直述，如 `列 "{token}" 不存在。表头包含：{表头列表}。` |
| `IMAGE_NOT_FOUND` | 图片路径不存在 | `图片文件不存在：{imagePath}。请检查路径。` |
| `UNSUPPORTED_IMAGE` | 图片格式不支持 | `不支持的图片格式：{ext}，仅支持 png / jpg / gif。` |
| `UNKNOWN` | 一切未预料异常 | `操作失败：{原始错误消息}`；文件解析失败专用：`无法解析文件：{filePath}，文件可能已损坏或不是有效的表格文件。` |

文案原则：**说人话、给下一步**——告诉用户/AI 发生了什么、可以怎么办，从不暴露堆栈和英文库错误。

---

## 4. 页面（触点）设计

### 4.0 为什么没有"页面"章节惯例的内容

本产品是 MCP 服务进程，**没有任何自己的窗口和页面**（需求文档第 6 节、架构文档第 4 节已明确），因此不存在"页面元素布局 / 页面跳转"。用户实际接触的只有三个触点：① npm 包说明页（README）；② AI 助手的对话窗口（不属于本系统，但本系统的输出在这里呈现）；③ 任务完成后的 Excel 成果文件。

本章按这三个触点描述"长什么样"，并把"正常 / 没数据 / 出错了"三种状态落到每个触点上。

### 4.1 触点一：npm 包说明页（README.md）

静态说明页，用户照着复制粘贴即可完成安装配置。`README.md` 的内容区块（自上而下，编码时照此结构写）：

1. **标题 + 一句话介绍**："mcp-excel —— 让 AI 助手直接读写你电脑上的 Excel 文件（支持 .xlsx / .xls / .csv）"。
2. **它能干什么**：功能 bullet 列表——读表问答、新建文件、改单元格、批量写入、公式、样式美化、排序/筛选/去重、插图。
3. **安装与配置**（最重要的区块，给可直接复制的代码块）：
   ````markdown
   在 AI 助手的 MCP 配置中加入：
   ```json
   {
     "mcpServers": {
       "excel": {
         "command": "npx",
         "args": ["-y", "mcp-excel"]
       }
     }
   }
   ```
   ````
   并注明：需要 Node.js 20 以上；保存配置后重启 AI 助手生效。
4. **使用示例**：3 条大白话对话示例（对应需求三场景）：
   - "帮我看看 D:\报表\2026年销售记录.xlsx 里'一月'那张表，总销售额是多少？"
   - "新建一个 Excel，12 张 Sheet，每张是一个月考勤表，第一行写姓名和 1 到 31 号。"
   - "把'费用表' B5 的 3000 改成 3500，最后加一行合计。"
5. **支持的格式与限制**（如实写明，防止错误预期）：
   - .xlsx 全功能；.xls 只读（修改请先另存为 .xlsx）；.csv 支持读写值，不支持多工作表/样式/图片。
   - 修改含图表、数据透视表、宏的文件可能丢失这些元素。
   - 写入的公式需在 Excel/WPS 中打开一次后才会显示计算结果。
6. **隐私说明**："所有处理都在你自己的电脑上完成，不上传任何数据。"
7. **License**：MIT。

三种状态对本触点不适用（静态页没有运行时状态），唯一要求是：**配置示例必须原样可用**，用户零修改粘贴即可跑通。

### 4.2 触点二：AI 对话窗口（本系统输出的呈现处）

对话窗口属于 AI 助手，但本工具的每一次输出都在这里被用户感知。三种状态：

**正常**——用户提问，AI 调用工具拿到 `success: true`，用人话回复：

```
用户：看看 D:\报表\2026年销售记录.xlsx "一月"那张表，总销售额多少？
  └─ AI 调用 read_range → { "success": true, "values": [["姓名","销售额",...],["张三",1200,...],...] }
AI：一月共 12 条销售记录，总销售额为 48,650 元。其中张三 12,300 元最高。
```

**没数据**——合法但为空的结果，`success` 仍为 `true`（空表/空区域/筛选无匹配都不是错误）：

```
用户：看看"二月"那张表。
  └─ AI 调用 read_range → { "success": true, "values": [] }
AI："二月"是一张空表，还没有任何内容。需要我先帮你写上表头吗？
```

```
用户：找出销售额超过 10 万的。
  └─ AI 调用 filter_range → { "success": true, "values": [["姓名","销售额","日期","备注"]] }
AI：没有符合条件的记录——一月最高的销售额是 12,300 元，没有超过 10 万的。
```

**出错了**——`success: false`，错误文案本身就是给最终用户看的人话（3.4 节），AI 原样转述并附上建议动作：

```
用户：把费用表 B5 改成 3500。
  └─ AI 调用 write_range → { "success": false, "error": "无法保存：C:\...\费用表.xlsx 正被其他程序（如 Excel）占用，请关闭后重试。" }
AI：没改成——费用表.xlsx 现在正被 Excel 打开着。请先在 Excel 里关闭这个文件，然后跟我说一声，我马上重试。
```

设计要点：工具文案已包含"发生了什么 + 下一步怎么办"，AI 不需要额外润色就能给出合格回复；`SHEET_NOT_FOUND` 等文案附带现有表名列表，AI 还可以据此自动纠正重试（如把"一月份"纠正为"一月"）。

### 4.3 触点三：成果 Excel 文件

任务完成后用户自己在 Excel/WPS 中打开的文件。三种状态：

**正常**：用户看到预期成果——新建的 12 张考勤 Sheet 及表头、改好的 B5 与合计行（公式显示计算结果）、锚定在 D2 的产品图、美化过的标题行。

**没数据**：合法的"空成果"——例如 `create_workbook` 刚执行完、尚未 `write_range` 时，文件里就是 12 张空 Sheet。这不是异常，是流程的中间态；AI 应告知用户"文件已建好，接下来要写什么内容？"。

**出错了**：由于原子写（全局约定 2），任何写操作失败时**原文件保持原样**——用户打开看到的是修改前的旧内容，不会遇到写了一半的损坏文件。这正是"出错状态"在成果文件上的表现：文件没变化 + 对话窗口里有明确的失败原因。临时文件 `<原名>.tmp-<pid>` 在任何路径下都会被清理（2.18 的 `finally`），不会在用户目录留下垃圾。

---

## 附：与架构文档的对应关系

- 模块 1（入口层）→ 2.5 `src/index.ts` + 2.6 `src/tools/index.ts`
- 模块 2（工具层）→ 2.7 ~ 2.13（7 个文件，15 个工具）
- 模块 3（引擎层）→ 2.14 ~ 2.18（3a 读取适配 = `reader.ts`，3b 编辑适配 = `writer.ts`）
- 架构第 5 节 v1 全部 15 个工具 → 本文档 2.8 ~ 2.13 全覆盖；架构明确暂缓的 `create_chart` / `merge_files` / `set_validation` / 模板 / PDF 导出**不在**本文件清单内，v1 不创建对应文件。
