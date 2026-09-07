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
  if (lower.includes("mk_e_unavailable") || lower.includes("0x800401e3") || lower.includes("getactiveobject")) {
    return "未检测到正在运行的 Excel，请先打开 Excel 并加载目标工作簿后重试。";
  }
  if (
    lower.includes("programmatic access") ||
    lower.includes("不信任") ||
    lower.includes("e_accessdenied") ||
    lower.includes("0x80070005") ||
    lower.includes("vba project")
  ) {
    return "无法访问 VBAProject。请在 Excel 中开启：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」。";
  }
  if (lower.includes("0x800ac472") || lower.includes("excel 正忙") || lower.includes("excel is busy")) {
    return "Excel 当前正忙，请稍后重试。";
  }
  if (lower.includes("rpc_e_disconnected") || lower.includes("0x80010108") || lower.includes("rpc server")) {
    return "Excel COM 调用失败，请检查 Excel 是否正常运行。";
  }
  if (
    (lower.includes("subscript out of range") || lower.includes("下标越界")) &&
    (lower.includes("worksheet") || lower.includes("worksheets") || lower.includes("sheet") || lower.includes("range"))
  ) {
    return "未找到目标工作表，请确认工作表名称是否正确（区分大小写）。";
  }
  if (lower.includes("cannot run the macro") || lower.includes("宏不可用") || lower.includes("无法运行宏")) {
    return "宏执行失败，请确认宏名称正确且工作簿中存在该宏。";
  }
  if (lower.includes("timeout") || lower.includes("timed out")) {
    return "操作超时，Excel 可能正在处理其他任务或弹出了对话框。请检查 Excel 窗口是否有弹窗需要处理，然后重试。";
  }
  if (lower.includes("being used by another") || lower.includes("正由另一个")) {
    return "文件正被其他程序占用，请关闭占用该文件的程序后重试。";
  }
  const normalized = raw
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  const preferred =
    normalized.find((line) => /exception calling|exception setting|cannot set|属性|property|error|失败|无法|未找到/i.test(line)) ||
    normalized.find((line) => !/^\+|^at\s|^at line:|^categoryinfo:|^fullyqualifiederrorid:/i.test(line)) ||
    "";
  const cleaned = preferred.replace(/^.*\.ps1\s*:\s*/i, "").trim();
  if (cleaned && !/^fullyqualifiederrorid\s*:/i.test(cleaned)) return cleaned;
  return raw.length > 1000 ? raw.substring(0, 1000) : raw;
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
      const fullOutput = [stdout.trim(), "--- STDERR ---", stderr.trim()].filter(Boolean).join("\n");
      return { success: false, message: friendlyError(stderr.trim()), output: fullOutput };
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
    const fullOutput = [stdout.trim(), stderr.trim(), fallbackMsg !== rawMsg ? fallbackMsg : ""].filter(Boolean).join("\n");
    return { success: false, message, output: fullOutput || undefined };
  }
}

/**
 * 生成获取 Excel.Application 对象的 PowerShell 脚本片段。
 * 如果提供了 excelProcessId，优先通过 PID 定位窗口并使用 AccessibleObjectFromWindow 获取 COM 对象；
 * 失败或未提供 PID 时回退到 GetActiveObject("Excel.Application")。
 */
export function buildExcelAttachScript(excelProcessId?: number): string {
  const fallbackLine = `$excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")`;
  if (!excelProcessId || excelProcessId <= 0) {
    return fallbackLine;
  }
  return `
$targetPid = ${excelProcessId}
$foundHwnd = [IntPtr]::Zero

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class JrExcelAttach {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

  [DllImport("oleacc.dll")]
  public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
}
"@

[JrExcelAttach]::EnumWindows({
  param($hWnd, $lParam)
  if (-not [JrExcelAttach]::IsWindowVisible($hWnd)) { return $true }
  $winPid = [uint32]0
  [void][JrExcelAttach]::GetWindowThreadProcessId($hWnd, [ref]$winPid)
  if ($winPid -ne $targetPid) { return $true }
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrExcelAttach]::GetClassName($hWnd, $sb, $sb.Capacity)
  if ($sb.ToString() -eq "XLMAIN") { $script:foundHwnd = $hWnd; return $false }
  return $true
}, [IntPtr]::Zero) | Out-Null

$excel = $null
if ($foundHwnd -ne [IntPtr]::Zero) {
  $guid = [Guid]::Parse("00020400-0000-0000-C000-000000000046")
  $obj = $null
  $hr = [JrExcelAttach]::AccessibleObjectFromWindow($foundHwnd, 0xFFFFFFF0, [ref]$guid, [ref]$obj)
  if ($hr -eq 0) {
    $excel = $obj.Application
  } else {
    Write-Warning "AccessibleObjectFromWindow failed, HRESULT=0x$($hr.ToString('X8'))"
  }
}
if ($excel -eq $null) {
  Write-Warning "通过 PID ${excelProcessId} attach Excel 失败，回退到 GetActiveObject。"
  ${fallbackLine}
}
if ($excel -eq $null) { throw "无法获取 Excel.Application" }
`;
}

/**
 * 检查 Excel 是否正在运行且已打开目标工作簿。
 * 如果提供了 excelProcessId，会优先 attach 到对应实例再检查工作簿。
 * 返回非 null 表示前置条件不满足，应直接返回该错误结果。
 */
export async function ensureExcelRunning(wbName: string, excelProcessId?: number): Promise<ExcelComResult | null> {
  const checkScript = `
$ErrorActionPreference = "Stop"
try {
    ${buildExcelAttachScript(excelProcessId)}
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
