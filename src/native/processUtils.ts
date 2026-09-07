/**
 * 进程检测工具
 *
 * 保留最小化的 OS 级进程检测逻辑，用于 Excel 进程跟踪与管理。
 * 本地 COM 调用仅通过生成的 PowerShell 脚本间接完成。
 */
import { exec } from "child_process";
import * as path from "path";
import { promisify } from "util";

const execAsync = promisify(exec);

function escapePowerShellSingleQuoted(value: string): string {
  return value.replace(/'/g, "''");
}

/** 检查指定 Excel 进程是否仍然存在 */
export async function isExcelProcessAlive(pid: number): Promise<boolean> {
  try {
    const result = await execAsync(
      `powershell -NoProfile -Command "try { $p = Get-Process -Id ${pid} -ErrorAction Stop; Write-Output 'ALIVE' } catch { Write-Output 'DEAD' }"`,
      { timeout: 10000 }
    );
    return (result.stdout || "").trim() === "ALIVE";
  } catch {
    return false;
  }
}

const sleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

/** 等待指定 PID 的进程退出，超时返回 false */
export async function waitForProcessExit(pid: number, timeoutMs = 5000): Promise<boolean> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (!(await isExcelProcessAlive(pid))) {
      return true;
    }
    await sleep(500);
  }
  return !(await isExcelProcessAlive(pid));
}

/** 强制终止指定进程及其子进程（Windows taskkill /T /F） */
export async function terminateProcessTree(pid: number): Promise<void> {
  await execAsync(`taskkill /PID ${pid} /T /F`, { timeout: 10000 });
}

/** 检查指定 PID 的进程是否仍存在任意窗口句柄 */
export async function hasAnyProcessWindow(pid: number): Promise<boolean> {
  try {
    const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class JrAnyWindowChecker {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern bool IsWindow(IntPtr hWnd);
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
$targetPid = ${pid}
$found = $false
[JrAnyWindowChecker]::EnumWindows({
  param($hWnd, $lParam)
  if (-not [JrAnyWindowChecker]::IsWindow($hWnd)) { return $true }
  $winPid = [uint32]0
  [void][JrAnyWindowChecker]::GetWindowThreadProcessId($hWnd, [ref]$winPid)
  if ($winPid -eq $targetPid) { $found = $true; return $false }
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($found) { Write-Output "PRESENT" } else { Write-Output "ABSENT" }
`;
    const result = await execAsync(
      `powershell -NoProfile -Command "${script.replace(/"/g, '\\"').replace(/\n/g, "; ")}"`,
      { timeout: 10000 }
    );
    return (result.stdout || "").trim() === "PRESENT";
  } catch {
    return false;
  }
}

/** 通过窗口标题反查包含目标工作簿的 Excel 进程 ID（不使用 COM） */
export async function findExcelProcessIdByWindow(workbookPath: string): Promise<number> {
  try {
    const path = await import("path");
    const wbName = path.basename(workbookPath);
    const wbNameNoExt = path.basename(workbookPath, path.extname(workbookPath));
    const script = `
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class JrExcelWindowFinder {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
  [DllImport("user32.dll")]
  public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")]
  public static extern bool IsIconic(IntPtr hWnd);
}
"@
$targetName = '${wbName.replace(/'/g, "''")}'
$targetNameNoExt = '${wbNameNoExt.replace(/'/g, "''")}'
$foundPid = 0
[JrExcelWindowFinder]::EnumWindows({
  param($hWnd, $lParam)
  $visible = [JrExcelWindowFinder]::IsWindowVisible($hWnd)
  $iconic = [JrExcelWindowFinder]::IsIconic($hWnd)
  if (-not $visible -and -not $iconic) { return $true }
  $sb = New-Object System.Text.StringBuilder 512
  [void][JrExcelWindowFinder]::GetWindowText($hWnd, $sb, 512)
  $title = $sb.ToString()
  if ($title -like "*$targetName*" -or $title -like "*$targetNameNoExt*") {
    $pid = [uint32]0
    [void][JrExcelWindowFinder]::GetWindowThreadProcessId($hWnd, [ref]$pid)
    $foundPid = [int]$pid
    return $false
  }
  return $true
}, [IntPtr]::Zero) | Out-Null
$foundPid
`;
    const result = await execAsync(
      `powershell -NoProfile -Command "${script.replace(/"/g, '\\"').replace(/\n/g, "; ")}"`,
      { timeout: 10000 }
    );
    const pid = parseInt((result.stdout || "").trim(), 10);
    return isNaN(pid) ? 0 : pid;
  } catch {
    return 0;
  }
}

/** 枚举所有 EXCEL.EXE 进程，返回 PID 与 CommandLine 列表 */
export async function listExcelProcesses(): Promise<Array<{ pid: number; commandLine: string }>> {
  try {
    const script = `
Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | Select-Object ProcessId, CommandLine | ForEach-Object {
  "$($_.ProcessId)|$($_.CommandLine)"
}
`;
    const encoded = Buffer.from(script, "utf16le").toString("base64");
    const result = await execAsync(
      `powershell -NoProfile -EncodedCommand ${encoded}`,
      { timeout: 10000 }
    );
    const processes: Array<{ pid: number; commandLine: string }> = [];
    for (const line of (result.stdout || "").split("\n")) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      const sepIndex = trimmed.indexOf("|");
      if (sepIndex === -1) continue;
      const pid = parseInt(trimmed.substring(0, sepIndex), 10);
      const commandLine = trimmed.substring(sepIndex + 1);
      if (!isNaN(pid)) {
        processes.push({ pid, commandLine });
      }
    }
    return processes;
  } catch {
    return [];
  }
}

/** 获取当前所有 EXCEL.EXE 进程的 PID 列表 */
export async function listExcelPids(): Promise<number[]> {
  const processes = await listExcelProcesses();
  return processes.map((p) => p.pid);
}

/**
 * 查找持有指定工作簿的 Excel 进程 PID。
 *
 * 结合以下信号按优先级匹配：
 * 1. WMI Win32_Process CommandLine 包含完整文件路径。
 * 2. 窗口标题包含工作簿文件名。
 * 3. excludePids 排除法（通常用于排除打开前已存在的 Excel 进程）。
 *
 * 多个 Excel 进程共存时，优先返回 CommandLine 唯一匹配；若 CommandLine 无法区分，
 * 则交叉验证窗口标题，仍失败时返回第一个 CommandLine 匹配项。
 */
export async function findExcelPidForWorkbook(path: string, excludePids?: number[]): Promise<number | undefined> {
  const excludeSet = new Set(excludePids || []);
  const normalizedPath = path.replace(/\//g, "\\");
  const lowerPath = normalizedPath.toLowerCase();

  const processes = await listExcelProcesses();
  const candidates = processes.filter((p) => !excludeSet.has(p.pid));

  // 信号 1：CommandLine 包含目标文件路径
  const commandLineMatches = candidates.filter((p) => {
    const cmd = (p.commandLine || "").toLowerCase();
    return cmd.includes(lowerPath);
  });

  if (commandLineMatches.length === 1) {
    return commandLineMatches[0].pid;
  }

  // 信号 2：窗口标题包含工作簿名（作为交叉验证或兜底）
  const windowPid = await findExcelProcessIdByWindow(path);
  const windowPidValid = windowPid !== 0 && !excludeSet.has(windowPid);

  if (commandLineMatches.length > 1) {
    // 多个 CommandLine 命中时，优先匹配窗口标题所在的进程
    const matchedByWindow = commandLineMatches.find((p) => p.pid === windowPid);
    if (matchedByWindow) {
      return matchedByWindow.pid;
    }
    // 窗口标题无法区分时，返回第一个候选（通常就是最新打开的实例）
    return commandLineMatches[0].pid;
  }

  if (windowPidValid) {
    return windowPid;
  }

  return undefined;
}

/**
 * 查找持有指定工作簿的 Excel 进程 PID（T6 冲突检测）。
 *
 * 优先使用 WMI Win32_Process CommandLine 匹配目标文件完整路径；
 * 多个进程命中时，用窗口标题交叉验证；无法区分时返回第一个候选。
 */
export async function findWorkbookOwnerPid(workbookPath: string): Promise<number | undefined> {
  const normalizedPath = workbookPath.replace(/\//g, "\\");
  const lowerPath = normalizedPath.toLowerCase();

  const processes = await listExcelProcesses();
  const matches = processes.filter((p) => {
    const cmd = (p.commandLine || "").toLowerCase();
    return cmd.includes(lowerPath);
  });

  if (matches.length === 1) {
    return matches[0].pid;
  }

  if (matches.length > 1) {
    const windowPid = await findExcelProcessIdByWindow(workbookPath);
    const matchedByWindow = matches.find((p) => p.pid === windowPid);
    if (matchedByWindow) {
      return matchedByWindow.pid;
    }
    return matches[0].pid;
  }

  return undefined;
}

/**
 * 通过 COM attach 到指定 PID 的 Excel 实例，保存并关闭目标工作簿，然后退出 Excel（T6 接管）。
 *
 * 优先使用 EnumWindows + AccessibleObjectFromWindow 按 PID 获取 Application；
 * 失败时回退到 GetActiveObject("Excel.Application")。
 * 返回是否成功完成保存、关闭、退出流程。
 */
export async function closeWorkbookAndQuit(pid: number, workbookPath: string): Promise<boolean> {
  const wbName = path.basename(workbookPath);
  const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class JrExcelAttachT6 {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
  [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
}
"@
$targetPid = ${pid}
$foundHwnd = [IntPtr]::Zero
[JrExcelAttachT6]::EnumWindows({
  param($hWnd, $lParam)
  if (-not [JrExcelAttachT6]::IsWindowVisible($hWnd)) { return $true }
  $winPid = [uint32]0
  [void][JrExcelAttachT6]::GetWindowThreadProcessId($hWnd, [ref]$winPid)
  if ($winPid -ne $targetPid) { return $true }
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrExcelAttachT6]::GetClassName($hWnd, $sb, $sb.Capacity)
  if ($sb.ToString() -eq "XLMAIN") { $script:foundHwnd = $hWnd; return $false }
  return $true
}, [IntPtr]::Zero) | Out-Null

$excel = $null
if ($foundHwnd -ne [IntPtr]::Zero) {
  $guid = [Guid]::Parse("00020400-0000-0000-C000-000000000046")
  $obj = $null
  $hr = [JrExcelAttachT6]::AccessibleObjectFromWindow($foundHwnd, 0xFFFFFFF0, [ref]$guid, [ref]$obj)
  if ($hr -eq 0) {
    $excel = $obj.Application
  }
}
if ($excel -eq $null) {
  $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
}
if ($excel -eq $null) { throw "无法获取 Excel.Application" }

$wb = $null
foreach ($w in $excel.Workbooks) {
  if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break }
}
if ($wb -ne $null) {
  $wb.Close($true)
}
$excel.Quit()
Write-Output "OK"
`;
  try {
    const encoded = Buffer.from(script, "utf16le").toString("base64");
    const result = await execAsync(
      `powershell -NoProfile -EncodedCommand ${encoded}`,
      { timeout: 30000 }
    );
    return (result.stdout || "").trim() === "OK";
  } catch {
    return false;
  }
}
