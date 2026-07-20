import { test, describe, before } from "node:test";
import assert from "node:assert";
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const rootDir = path.resolve(fileURLToPath(import.meta.url), "..", "..");
const pkgJson = JSON.parse(
  fs.readFileSync(path.join(rootDir, "package.json"), "utf8")
);

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

function exec(
  command: string,
  args: string[],
  cwd: string
): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd,
      shell: true,
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout?.on("data", (d: Buffer) => (stdout += d.toString()));
    child.stderr?.on("data", (d: Buffer) => (stderr += d.toString()));
    child.on("close", (code) => resolve({ code, stdout, stderr }));
    child.on("error", reject);
  });
}

describe("配置 / 发布文件", () => {
  before(async () => {
    await build();
  });

  test("CF-01 构建通过", () => {
    assert.ok(
      fs.existsSync(path.join(rootDir, "dist", "index.js")),
      "dist/index.js 未生成"
    );
  });

  test("CF-02 bin 入口可启动", async () => {
    const child = spawn("node", [path.join(rootDir, "dist", "index.js")], {
      cwd: rootDir,
      stdio: ["pipe", "pipe", "pipe"],
    });
    let stderr = "";
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", (d: string) => (stderr += d));

    // 最多等 2 秒，直到 stderr 出现启动日志
    await new Promise<void>((resolve) => {
      const timer = setInterval(() => {
        if (stderr.includes("已启动")) {
          clearInterval(timer);
          resolve();
        }
      }, 50);
      setTimeout(() => {
        clearInterval(timer);
        resolve();
      }, 2000);
    });

    child.stdin.end();
    child.kill();

    await new Promise((resolve) => {
      child.on("close", resolve);
      setTimeout(() => resolve(undefined), 1000);
    });

    assert.ok(stderr.includes("已启动"), `bin 入口应输出启动日志，实际：${stderr}`);
  });

  test("CF-03 发布内容最小化", async () => {
    const { stdout, stderr } = await exec("npm", ["pack", "--dry-run"], rootDir);
    const output = stdout + stderr;
    assert.ok(output.includes("Tarball Contents"), "npm pack 输出异常");

    const lines = output.split(/\r?\n/);
    const files: string[] = [];
    let inContents = false;
    for (const line of lines) {
      if (line.includes("Tarball Contents")) {
        inContents = true;
        continue;
      }
      if (line.includes("Tarball Details")) break;
      if (inContents) {
        const trimmed = line.replace(/^npm notice\s+/, "").trim();
        // 格式："2.2kB README.md"，取空格后文件名
        const fileName = trimmed.replace(/^\S+\s+/, "").trim();
        if (fileName) files.push(fileName);
      }
    }

    assert.ok(
      files.includes("dist/index.js"),
      `tarball 应包含 dist/index.js，实际：${files.join(", ")}`
    );
    assert.ok(
      files.includes("package.json"),
      `tarball 应包含 package.json，实际：${files.join(", ")}`
    );
    assert.ok(
      files.includes("README.md"),
      `tarball 应包含 README.md，实际：${files.join(", ")}`
    );

    const forbidden = files.some(
      (f) =>
        f.startsWith("src/") ||
        f.startsWith("tests/") ||
        f.startsWith("node_modules/")
    );
    assert.ok(!forbidden, `tarball 不应包含 src/ tests/ node_modules/：${files.join(", ")}`);
  });

  test("CF-04 engines 声明", () => {
    assert.ok(pkgJson.engines, "package.json 缺少 engines");
    assert.ok(
      pkgJson.engines.node.startsWith(">=") && Number(pkgJson.engines.node.slice(2)) >= 20,
      `engines.node 应声明 >=20，实际：${pkgJson.engines.node}`
    );
  });

  test("CF-05 .gitignore 生效", async () => {
    const { stdout } = await exec("git", ["status", "--porcelain"], rootDir);
    const lines = stdout.split(/\r?\n/).filter((l) => l.trim());
    for (const line of lines) {
      const p = line.slice(3).trim();
      assert.ok(
        !p.startsWith("dist/") && !p.includes("node_modules/"),
        `.gitignore 未生效：${line}`
      );
    }
  });

  test("CF-06 README 配置示例可用", () => {
    const readme = fs.readFileSync(path.join(rootDir, "README.md"), "utf8");
    const match = readme.match(/```json\s*([\s\S]*?)\s*```/);
    assert.ok(match, "README 缺少 JSON 配置块");

    let cfg: any;
    assert.doesNotThrow(() => {
      cfg = JSON.parse(match![1]);
    }, `README 中的 JSON 配置块无法解析：${match![1]}`);

    assert.strictEqual(cfg.mcpServers?.excel?.command, "npx");
    assert.deepStrictEqual(cfg.mcpServers?.excel?.args, ["-y", "mcp-excel"]);

    const lower = readme.toLowerCase();
    assert.ok(
      lower.includes("node.js 20") || lower.includes("node 20"),
      "README 应声明 Node 20 要求"
    );
  });

  test("CF-07 README 如实声明限制", () => {
    const readme = fs.readFileSync(path.join(rootDir, "README.md"), "utf8");
    assert.ok(readme.includes("图表"), "README 应声明图表可能丢失");
    assert.ok(
      readme.includes("数据透视表") || readme.includes("透视表"),
      "README 应声明数据透视表可能丢失"
    );
    assert.ok(
      readme.includes("宏") || readme.includes("VBA"),
      "README 应声明宏/VBA 可能丢失"
    );
    assert.ok(readme.includes("公式"), "README 应声明公式缓存限制");
    assert.ok(
      readme.includes("隐私") && readme.includes("不上传"),
      "README 应声明隐私说明"
    );
  });
});
