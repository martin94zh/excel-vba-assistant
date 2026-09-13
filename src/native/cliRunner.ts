/**
 * Excel VBA Assistant - excelcli 运行器
 *
 * ExcelMcp 2.x 的 CLI（excelcli.exe）通过共享后台 daemon（命名管道服务）持有 Excel：
 * 所有 excelcli 进程都连接同一个 daemon，会话跨进程持久。
 * 这解决了 MCP stdio 模式下"每个客户端各起一个 server、会话进程内私有"的隔离问题——
 * 插件打开的工作簿，AI 用任意新起的 excelcli 进程即可操作同一会话。
 */
import { spawn } from "child_process";
import { existsSync } from "fs";
import { join } from "path";

const CLI_EXE_NAME = "excelcli.exe";

/** 在插件 dist 目录中定位 excelcli.exe */
export function findLocalExcelCli(distDir: string): string | undefined {
  const candidate = join(distDir, CLI_EXE_NAME);
  if (existsSync(candidate)) return candidate;
  return undefined;
}

export interface ExcelCliResult {
  success: boolean;
  /** -q 模式下解析出的 JSON 对象（若有） */
  data?: Record<string, unknown>;
  /** 原始 stdout/stderr */
  raw: string;
  error?: string;
}

/**
 * 执行 excelcli 命令并解析输出。
 * -q 模式输出单行 JSON（如 {"success":true,"sessionId":"..."}）；
 * 兼容多行输出：取最后一条可解析的 JSON 行。
 */
export async function runExcelCli(exePath: string, args: string[], timeoutMs = 90000): Promise<ExcelCliResult> {
  return new Promise((resolve) => {
    const child = spawn(exePath, args, { stdio: ["ignore", "pipe", "pipe"], windowsHide: true });
    let stdout = "";
    let stderr = "";
    let settled = false;

    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill();
      resolve({ success: false, raw: stdout, error: `excelcli ${args.join(" ")} 超时（${timeoutMs}ms）` });
    }, timeoutMs);

    child.stdout.on("data", (d: Buffer) => (stdout += d.toString()));
    child.stderr.on("data", (d: Buffer) => (stderr += d.toString()));

    child.on("error", (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ success: false, raw: stdout, error: `无法启动 excelcli：${err.message}` });
    });

    child.on("close", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      const output = (stdout || "").trim();
      let data: Record<string, unknown> | undefined;
      for (const line of output.split("\n").reverse()) {
        const t = line.trim();
        if (!t) continue;
        try {
          data = JSON.parse(t) as Record<string, unknown>;
          break;
        } catch {
          // 继续尝试上一行
        }
      }
      const logicalFail = data?.success === false || data?.isError === true;
      const noDataFail = data === undefined && code !== 0;
      resolve({
        success: !logicalFail && !noDataFail,
        data,
        raw: output || stderr.trim(),
        error: logicalFail
          ? String(data?.error ?? data?.errorMessage ?? "excelcli 命令执行失败")
          : noDataFail
            ? stderr.trim() || `excelcli 退出码 ${code}`
            : undefined,
      });
    });
  });
}
