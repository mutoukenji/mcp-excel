import { test, describe, before } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const rootDir = path.resolve(fileURLToPath(import.meta.url), "..", "..");
const srcIndex = path.join(rootDir, "src", "index.ts");
const distIndex = path.join(rootDir, "dist", "index.js");

async function build(): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn("npm", ["run", "build"], {
      cwd: rootDir,
      shell: true,
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout?.on("data", (d: Buffer) => (stdout += d.toString()));
    child.stderr?.on("data", (d: Buffer) => (stderr += d.toString()));
    child.on("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`build failed (code ${code}):\n${stderr}\n${stdout}`));
    });
    child.on("error", reject);
  });
}

function walkSrcFiles(dir: string, out: string[]) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) walkSrcFiles(p, out);
    else if (entry.name.endsWith(".ts")) out.push(p);
  }
}

function firstLine(filePath: string): string {
  return fs.readFileSync(filePath, "utf8").split(/\r?\n/)[0];
}

describe("src/index.ts", () => {
  before(async () => {
    await build();
  });

  test("IX-01 shebang 保留", () => {
    assert.strictEqual(firstLine(srcIndex), "#!/usr/bin/env node");
    assert.strictEqual(firstLine(distIndex), "#!/usr/bin/env node");
  });

  test("IX-02 stdout 纯净", async () => {
    const srcFiles: string[] = [];
    walkSrcFiles(path.join(rootDir, "src"), srcFiles);
    for (const p of srcFiles) {
      const content = fs.readFileSync(p, "utf8");
      assert.ok(
        !content.includes("console.log"),
        `src 文件 ${path.relative(rootDir, p)} 包含 console.log`
      );
      assert.ok(
        !content.includes("process.stdout.write"),
        `src 文件 ${path.relative(rootDir, p)} 包含 process.stdout.write`
      );
    }

    const child = spawn("node", [distIndex], {
      cwd: rootDir,
      stdio: ["pipe", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (d: string) => (stdout += d));
    child.stderr.on("data", (d: string) => (stderr += d));

    const init =
      JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "test", version: "1.0.0" },
        },
      }) + "\n";
    child.stdin.write(init);

    await new Promise((resolve) => setTimeout(resolve, 800));

    child.stdin.end();
    child.kill();

    await new Promise((resolve) => {
      child.on("close", resolve);
      setTimeout(() => resolve(undefined), 1000);
    });

    assert.ok(stderr.includes("已启动"), `stderr 应包含启动提示，实际：${stderr}`);

    const lines = stdout.split(/\r?\n/).filter((l) => l.trim());
    assert.ok(lines.length > 0, "initialize 后应收到 JSON-RPC 响应");
    for (const line of lines) {
      let parsed: unknown;
      try {
        parsed = JSON.parse(line);
      } catch (e) {
        assert.fail(`stdout 包含非 JSON 行：${line}`);
      }
      assert.strictEqual(
        (parsed as Record<string, unknown>).jsonrpc,
        "2.0",
        `stdout 非法 JSON-RPC 行：${line}`
      );
    }
  });

  test("IX-03 进程启动并握手 / IX-04 启动日志走 stderr", async () => {
    const transport = new StdioClientTransport({
      command: "node",
      args: [distIndex],
      cwd: rootDir,
      stderr: "pipe",
    });

    let stderr = "";
    const errStream = transport.stderr;
    if (errStream) {
      errStream.setEncoding("utf8");
      errStream.on("data", (d: string) => (stderr += d));
    }

    const client = new Client({ name: "test-client", version: "1.0.0" });
    await client.connect(transport);

    try {
      const tools = await client.listTools();
      assert.strictEqual(tools.tools.length, 15, `应注册 15 个工具，实际 ${tools.tools.length}`);

      for (const t of tools.tools) {
        assert.ok(t.title, `${t.name} 缺少 title`);
        assert.ok(t.description, `${t.name} 缺少 description`);
        const text = (t.title ?? "") + (t.description ?? "");
        assert.ok(/[\u4e00-\u9fa5]/.test(text), `${t.name} 的 title/description 缺少中文`);
      }

      const tmpDir = await fs.promises.mkdtemp(
        path.join(os.tmpdir(), "mcp-excel-ix-")
      );
      const filePath = path.join(tmpDir, "e2e.xlsx");
      const res = await client.callTool({
        name: "create_workbook",
        arguments: { filePath, sheets: ["Sheet1"] },
      });
      const text = (res.content[0] as { text: string }).text;
      const parsed = JSON.parse(text);
      assert.strictEqual(parsed.success, true, `create_workbook 失败：${text}`);
      assert.strictEqual(parsed.sheetsCreated, 1);

      await fs.promises.rm(tmpDir, { recursive: true, force: true });
    } finally {
      await client.close();
    }

    assert.ok(stderr.includes("已启动"), `stderr 应包含启动提示，实际：${stderr}`);
  });
});
