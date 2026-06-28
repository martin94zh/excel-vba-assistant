/**
 * Excel VBA Assistant - PowerShell 运行时
 *
 * 通过生成 PowerShell 脚本调用 Excel COM Automation，
 * 实现对已打开 Excel 工作簿 / VBAProject 的读写操作。
 *
 * 该模块是本地服务层的替代实现：
 * 任务文档中描述的 C# .NET + ASP.NET Core 服务在本版本中以
 * TypeScript + PowerShell COM 的方式实现，架构更轻量、无需额外进程。
 */
import { exec } from "child_process";
import { randomBytes } from "crypto";
import { writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { promisify } from "util";

const execAsync = promisify(exec);

/** COM 调用结果 */
export interface ExcelComResult {
  success: boolean;
  message: string;
  output?: string;
  details?: Record<string, unknown>;
}

/** 将原始错误信息映射为中文友好提示 */
export function friendlyError(raw: string): string {
  const lower = raw.toLowerCase();
  const hintLines: string[] = [];

  if (lower.includes("mk_e_unavailable") || lower.includes("0x800401e3") || lower.includes("getactiveobject")) {
    hintLines.push("未检测到正在运行的 Excel，请先打开 Excel 并加载目标工作簿后重试。");
  }
  if (
    lower.includes("programmatic access") ||
    lower.includes("不信任") ||
    lower.includes("e_accessdenied") ||
    lower.includes("0x80070005") ||
    lower.includes("vba project")
  ) {
    hintLines.push("无法访问 VBAProject。请在 Excel 中开启：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」。");
  }
  if (lower.includes("0x800ac472") || lower.includes("excel 正忙") || lower.includes("excel is busy")) {
    hintLines.push("Excel 当前正忙，请稍后重试。");
  }
  if (lower.includes("rpc_e_disconnected") || lower.includes("0x80010108") || lower.includes("rpc server")) {
    hintLines.push("Excel COM 调用失败，请检查 Excel 是否正常运行。");
  }
  if (
    (lower.includes("subscript out of range") || lower.includes("下标越界")) &&
    (lower.includes("worksheet") || lower.includes("worksheets") || lower.includes("sheet") || lower.includes("range"))
  ) {
    hintLines.push("未找到目标工作表，请确认工作表名称是否正确（区分大小写）。");
  }
  if (lower.includes("cannot run the macro") || lower.includes("宏不可用") || lower.includes("无法运行宏")) {
    hintLines.push("宏执行失败，请确认宏名称正确且工作簿中存在该宏。");
  }
  if (lower.includes("timeout") || lower.includes("timed out")) {
    hintLines.push("操作超时，Excel 可能正在处理其他任务或弹出了对话框。请检查 Excel 窗口是否有弹窗需要处理，然后重试。");
  }
  if (lower.includes("being used by another") || lower.includes("正由另一个")) {
    hintLines.push("文件正被其他程序占用，请关闭占用该文件的程序后重试。");
  }

  // 清理 PowerShell 元信息，保留原始错误上下文
  const cleanedRaw = raw
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0 && !/^#< clixml|^<objs |xmlns="http:\/\/schemas\.microsoft\.com\/powershell\/2004\/04"/i.test(line))
    .join("\n");

  const detail = cleanedRaw.length > 1500 ? cleanedRaw.substring(0, 1500) + "\n..." : cleanedRaw;

  if (hintLines.length > 0) {
    return hintLines.join("\n") + "\n原始错误：\n" + detail;
  }
  return detail || raw;
}

/** 转义 PowerShell 单引号字符串 */
export function escapePowerShellSingleQuoted(value: string): string {
  return value.replace(/'/g, "''");
}

export function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * 执行 PowerShell 脚本。
 * 脚本写入临时 .ps1 文件后通过 powershell.exe 执行，
 * 强制 UTF-8 编码以正确处理中文输出。
 */
export async function runPowerShell(script: string, timeoutMs = 120000): Promise<ExcelComResult> {
  const tmpFile = join(tmpdir(), `excel_vba_${randomBytes(8).toString("hex")}.ps1`);
  const bom = "\uFEFF";
  const utf8Header = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
`;
  await writeFile(tmpFile, bom + utf8Header + script, "utf-8");
  try {
    const { stdout, stderr } = await execAsync(
      `chcp 65001 >nul & powershell -NoProfile -ExecutionPolicy Bypass -File "${tmpFile}"`,
      { timeout: timeoutMs, maxBuffer: 10 * 1024 * 1024, encoding: "utf-8" }
    );
    execAsync(`del "${tmpFile}"`).catch(() => {});
    if (stderr && stderr.trim()) {
      return { success: false, message: friendlyError(stderr.trim()), output: stdout };
    }
    return { success: true, message: "执行成功", output: stdout.trim() };
  } catch (err: any) {
    execAsync(`del "${tmpFile}"`).catch(() => {});
    const stderr = err?.stderr?.trim?.() || "";
    const stdout = err?.stdout?.trim?.() || "";
    const fallbackMsg = err instanceof Error ? err.message : String(err);
    const rawMsg = stderr || stdout || fallbackMsg;
    let message = friendlyError(rawMsg);
    if (!stderr && !stdout && message === fallbackMsg && err?.code !== undefined) {
      message = `${message} (exit code: ${String(err.code)})`;
    }
    return { success: false, message, output: stdout || undefined };
  }
}

/**
 * 检查 Excel 是否正在运行且已打开目标工作簿。
 * 返回非 null 表示前置条件不满足，应直接返回该错误结果。
 */
export async function ensureExcelRunning(wbName: string): Promise<ExcelComResult | null> {
  const checkScript = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $found = $false
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $found = $true; break }
    }
    if (-not $found) {
        Write-Output "WB_NOT_FOUND"
    } else {
        Write-Output "OK"
    }
} catch {
    Write-Output "EXCEL_NOT_RUNNING"
}
`;
  const result = await runPowerShell(checkScript);
  if (!result.success) return null;
  const output = (result.output || "").trim();
  if (output === "EXCEL_NOT_RUNNING") {
    return { success: false, message: "未检测到正在运行的 Excel。" };
  }
  if (output === "WB_NOT_FOUND") {
    return { success: false, message: `Excel 已运行，但未找到文件「${wbName}」。请在 Excel 中打开该文件后重试。` };
  }
  return null;
}
