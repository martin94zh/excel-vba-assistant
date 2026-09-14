/**
 * Excel VBA 宏运行器 - 支持弹窗检测与自动点击
 *
 * 移植并简化自原项目 vbe-macro.ts 的实现：
 * 运行宏时通过子进程异步执行，同时轮询检测 Excel/VBA 弹出的 MsgBox、
 * InputBox、运行时错误等对话框，并自动点击常见按钮（确定/结束/取消等）。
 */

import { randomBytes } from "crypto";
import { unlink, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { basename, join } from "path";
import { spawn } from "child_process";

import {
  buildExcelAttachScript,
  ensureExcelRunning,
  escapePowerShellSingleQuoted,
  friendlyError,
  runPowerShell,
} from "../runtime/powershell";
import type { ExcelComResult } from "../runtime/powershell";

export interface MacroDialogInfo {
  handle: string;
  processId: number;
  processName?: string;
  title: string;
  className: string;
  text: string;
  buttons: string[];
  width?: number;
  height?: number;
  kind?: "vb_runtime_error" | "error" | "confirmation" | "info" | "unknown";
  autoHandled?: boolean;
  autoAction?: string;
  /** 弹窗中是否含输入框（InputBox 类） */
  hasEdit?: boolean;
}

/** 宏运行时的交互弹窗处理策略（宏运行桥的 AI 测试场景使用全自动模式） */
export interface MacroDialogPolicy {
  /** auto = 弹窗全自动处理（MsgBox 点确定、确认框点 confirmButton、InputBox 填 inputValue）；errors = 仅自动处理错误弹窗（默认） */
  mode?: "auto" | "errors";
  /** 确认（是/否）对话框点击的按钮，默认"取消"（安全值） */
  confirmButton?: string;
  /** InputBox 自动填入的文本（填入后提交） */
  inputValue?: string;
}

interface MacroRunnerStatus {
  state: "success" | "error";
  message: string;
  details?: Record<string, unknown>;
  timestamp?: string;
}

const MACRO_MONITOR_TIMEOUT_MS = 45000;
const MACRO_DIALOG_POLL_MS = 300;
const MACRO_DIALOG_STABLE_MS = 450;
const MACRO_RUNNER_STATUS_GRACE_MS = 1200;
const MACRO_RUNNER_STATUS_POLL_MS = 100;
const SUPPORTED_DIALOG_PROCESS_PATTERN = "^(?i)(excel|et|wps)$";
const SUPPORTED_DIALOG_TITLE_PATTERN = "(?i)(microsoft excel|microsoft visual basic|excel|visual basic|wps|kingsoft|et|调用堆栈|call stack)";

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function readJsonFile<T>(filePath: string): Promise<T | null> {
  try {
    const { readFile } = await import("fs/promises");
    const raw = await readFile(filePath, "utf-8");
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

function classifyMacroDialog(dialog: MacroDialogInfo): MacroDialogInfo["kind"] {
  const cleanTitle = dialog.title.trim();
  const body = `${dialog.title}\n${dialog.text}`;
  // VBA 错误弹窗的标题恰好是 "Microsoft Visual Basic [for Applications]"（无后缀）；
  // VBE 主窗口/代码窗口标题带 " - 文件 [中断] - 模块" 等后缀，不能凭包含关系误判
  if (/^microsoft visual basic( for applications)?$/i.test(cleanTitle) || /运行时错误|编译错误|compile error/i.test(body)) {
    return "vb_runtime_error";
  }
  const haystack = `${cleanTitle}\n${dialog.text}\n${dialog.buttons.join(" ")}`.toLowerCase();
  if (
    haystack.includes("error")
    || haystack.includes("错误")
    || haystack.includes("exception")
    || haystack.includes("failed")
    || haystack.includes("cannot")
    || haystack.includes("无法")
    || haystack.includes("不能")
  ) {
    return "error";
  }
  const hasYesNo = /(yes|no|是|否)/i.test(haystack);
  const hasCancel = /(cancel|取消)/i.test(haystack);
  if (hasYesNo || hasCancel || dialog.buttons.length >= 2) {
    return "confirmation";
  }
  if (dialog.buttons.length >= 1) {
    return "info";
  }
  return "unknown";
}

function summarizeDialog(dialog: MacroDialogInfo): string {
  const title = dialog.title || "未命名弹窗";
  const text = dialog.text.replace(/\s+/g, " ").trim();
  const shortText = text.length > 120 ? `${text.slice(0, 120)}...` : text;
  return shortText ? `${title}: ${shortText}` : title;
}

function buildDialogFingerprint(dialog: Pick<MacroDialogInfo, "processName" | "title" | "text" | "buttons">): string {
  return [
    (dialog.processName || "").trim().toLowerCase(),
    dialog.title.trim(),
    dialog.text.replace(/\s+/g, " ").trim(),
    dialog.buttons.join(","),
  ].join("|");
}

function pickDialogAction(dialog: MacroDialogInfo): { buttons: string[]; actionLabel: string; terminal: boolean } | null {
  const haystack = `${dialog.title}\n${dialog.text}\n${dialog.buttons.join(" ")}`;
  // 调用堆栈窗口：中断模式的伴生窗口，直接关闭
  if (/调用堆栈|call stack/i.test(dialog.title)) {
    return { buttons: ["关闭", "Close"], actionLabel: "关闭", terminal: false };
  }
  // VBE 中断模式下修改代码时的系统恢复确认框，点"是"即重置工程（解除阻塞的正确恢复动作）
  if (/重新设置工程|reset the project/i.test(haystack)) {
    return { buttons: ["是", "Yes", "确定", "OK"], actionLabel: "是", terminal: true };
  }
  const kind = dialog.kind || classifyMacroDialog(dialog);
  switch (kind) {
    case "vb_runtime_error":
      return { buttons: ["结束", "End", "确定", "OK"], actionLabel: "结束", terminal: true };
    case "error":
      return { buttons: ["确定", "OK", "关闭", "Close"], actionLabel: "确定", terminal: true };
    case "info":
      return { buttons: ["确定", "OK", "继续", "Continue"], actionLabel: "确定", terminal: false };
    case "confirmation":
      return { buttons: ["取消", "Cancel", "否", "No", "关闭", "Close"], actionLabel: "取消", terminal: true };
    default:
      return null;
  }
}

function resolveDialogAction(
  dialog: MacroDialogInfo,
  action?: string,
  buttonText?: string
): { buttons: string[]; actionLabel: string; terminal: boolean } | null {
  if (buttonText && buttonText.trim()) {
    return { buttons: [buttonText.trim()], actionLabel: buttonText.trim(), terminal: false };
  }
  if (!action || action.trim().toLowerCase() === "auto") {
    return pickDialogAction(dialog);
  }
  const normalized = action.trim().toLowerCase();
  switch (normalized) {
    case "ok":
    case "confirm":
    case "确定":
      return { buttons: ["确定", "OK", "继续", "Continue", "重试", "Retry"], actionLabel: "确定", terminal: false };
    case "cancel":
    case "取消":
      return { buttons: ["取消", "Cancel"], actionLabel: "取消", terminal: true };
    case "yes":
    case "是":
      return { buttons: ["是", "Yes"], actionLabel: "是", terminal: false };
    case "no":
    case "否":
      return { buttons: ["否", "No"], actionLabel: "否", terminal: true };
    case "end":
    case "结束":
      return { buttons: ["结束", "End"], actionLabel: "结束", terminal: true };
    case "close":
    case "关闭":
      return { buttons: ["关闭", "Close"], actionLabel: "关闭", terminal: true };
    case "ignore":
    case "忽略":
      return { buttons: ["忽略", "Ignore"], actionLabel: "忽略", terminal: false };
    case "retry":
    case "重试":
      return { buttons: ["重试", "Retry"], actionLabel: "重试", terminal: false };
    default:
      return null;
  }
}

function buildMacroRunnerScript(params: {
  filePath: string;
  workbookName: string;
  macroName: string;
  statusPath: string;
  excelProcessId?: number;
}): string {
  const absPath = params.filePath.replace(/\//g, "\\");
  return `
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$statusPath = '${escapePowerShellSingleQuoted(params.statusPath.replace(/\//g, "\\"))}'
$macroName = '${escapePowerShellSingleQuoted(params.macroName.trim())}'
function Write-RunnerStatus([string]$state, [string]$message, [hashtable]$details) {
  $payload = [ordered]@{
    state = $state
    message = $message
    details = $details
    timestamp = (Get-Date).ToString("o")
  }
  $json = $payload | ConvertTo-Json -Depth 8
  [System.IO.File]::WriteAllText($statusPath, $json, [System.Text.UTF8Encoding]::new($false))
}
function Get-WorkbookQualifiedMacroName([string]$workbookName, [string]$macroPath) {
  $safeWorkbookName = $workbookName -replace "'", "''"
  return "'" + $safeWorkbookName + "'!" + $macroPath
}
function Resolve-ExcelRunFormat($workbook, [string]$requestedMacroName) {
  $macroPath = $requestedMacroName.Trim()
  if ($macroPath.Contains("!")) { return $macroPath }

  $projectName = ""
  try { $projectName = [string]$workbook.VBProject.Name } catch {}
  $defaultProjectPrefix = "VBAProject."
  if (-not [string]::IsNullOrWhiteSpace($projectName) -and $macroPath.StartsWith($projectName + ".", [System.StringComparison]::OrdinalIgnoreCase)) {
    $macroPath = $macroPath.Substring($projectName.Length + 1)
  } elseif ($macroPath.StartsWith($defaultProjectPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    $macroPath = $macroPath.Substring($defaultProjectPrefix.Length)
  }

  return (Get-WorkbookQualifiedMacroName ([string]$workbook.Name) $macroPath)
}
$excel = $null
$oldSecurity = $null
try {
  ${buildExcelAttachScript(params.excelProcessId)}
  $wb = $null
  foreach ($w in $excel.Workbooks) {
    if ($w.Name -eq '${escapePowerShellSingleQuoted(params.workbookName)}') { $wb = $w; break }
  }
  if ($wb -eq $null) {
    $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(absPath)}')
  }
  $oldSecurity = $excel.AutomationSecurity
  $excel.AutomationSecurity = 1
  try {
    $vbe = $excel.VBE
    $null = $vbe.ActiveVBProject.Name
  } catch {}
  Start-Sleep -Milliseconds 250
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.ScreenUpdating = $false
  $runFormat = Resolve-ExcelRunFormat $wb $macroName
  $excel.Run($runFormat)
  Write-RunnerStatus "success" "宏 ${escapePowerShellSingleQuoted(params.macroName.trim())} 已在 ${escapePowerShellSingleQuoted(params.workbookName)} 中执行完成" @{
    macroName = '${escapePowerShellSingleQuoted(params.macroName.trim())}'
    workbook = '${escapePowerShellSingleQuoted(params.workbookName)}'
    runFormat = $runFormat
  }
} catch {
  $err = $_
  $details = @{
    macroName = '${escapePowerShellSingleQuoted(params.macroName.trim())}'
    workbook = '${escapePowerShellSingleQuoted(params.workbookName)}'
    rawError = [string]$err.Exception.Message
    exceptionType = if ($err.Exception) { $err.Exception.GetType().FullName } else { "" }
    hresult = if ($err.Exception) { ('0x{0:X8}' -f ($err.Exception.HResult -band 0xffffffff)) } else { "" }
    runFormat = if ($runFormat) { $runFormat } else { "" }
  }
  Write-RunnerStatus "error" $details.rawError $details
  exit 1
} finally {
  if ($null -ne $excel -and $null -ne $oldSecurity) {
    try { $excel.AutomationSecurity = $oldSecurity } catch {}
  }
}
`;
}

async function waitForRunnerStatus(statusPath: string, waitMs: number): Promise<MacroRunnerStatus | null> {
  const deadline = Date.now() + waitMs;
  while (true) {
    const status = await readJsonFile<MacroRunnerStatus>(statusPath);
    if (status) return status;
    if (Date.now() >= deadline) return null;
    await sleep(MACRO_RUNNER_STATUS_POLL_MS);
  }
}

async function inspectExcelDialogs(): Promise<{ dialogs: MacroDialogInfo[]; error?: string; rawOutput?: string }> {
  const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class VbeWin32 {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

  [StructLayout(LayoutKind.Sequential)]
  public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumChildWindows(IntPtr hWnd, EnumChildProc callback, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowTextLength(IntPtr hWnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);

  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
}
"@
function Get-WindowTextSafe([IntPtr]$hWnd) {
  $len = [VbeWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][VbeWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][VbeWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ProcessNameSafe([int]$targetProcessId) {
  try { return (Get-Process -Id $targetProcessId -ErrorAction Stop).ProcessName } catch { return "" }
}
$results = New-Object System.Collections.Generic.List[object]
[VbeWin32]::EnumWindows({
  param($hWnd, $lParam)
  if (-not [VbeWin32]::IsWindowVisible($hWnd)) { return $true }
  $windowProcessId = 0; [void][VbeWin32]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
  $title = Get-WindowTextSafe $hWnd
  $className = Get-ClassNameSafe $hWnd
  $processName = Get-ProcessNameSafe $windowProcessId
  if ($processName -notmatch '${SUPPORTED_DIALOG_PROCESS_PATTERN}') { return $true }
  $buttons = New-Object System.Collections.Generic.List[string]
  $textParts = New-Object System.Collections.Generic.List[string]
  $script:hasEdit = $false
  [VbeWin32]::EnumChildWindows($hWnd, {
    param($childHwnd, $childLparam)
    if (-not [VbeWin32]::IsWindowVisible($childHwnd)) { return $true }
    $childClass = Get-ClassNameSafe $childHwnd
    # 输入框类控件先于空文本判断：InputBox 的 Edit 初始内容为空，不能因空文本跳过
    if ($childClass -in @("Edit", "RichEdit20W", "RichEdit50W", "RICHEDIT50W")) {
      $script:hasEdit = $true
      $childText2 = Get-WindowTextSafe $childHwnd
      if (-not [string]::IsNullOrWhiteSpace($childText2) -and -not $textParts.Contains($childText2)) { $textParts.Add($childText2) | Out-Null }
      return $true
    }
    $childText = Get-WindowTextSafe $childHwnd
    if ([string]::IsNullOrWhiteSpace($childText)) { return $true }
    if ($childClass -eq "Button") {
      if (-not $buttons.Contains($childText)) { $buttons.Add($childText) | Out-Null }
      return $true
    }
    if ($childClass -eq "Static") {
      if (-not $textParts.Contains($childText)) { $textParts.Add($childText) | Out-Null }
    }
    return $true
  }, [IntPtr]::Zero) | Out-Null

  if ($buttons.Count -eq 0) { return $true }

  $rect = New-Object VbeWin32+RECT
  [void][VbeWin32]::GetWindowRect($hWnd, [ref]$rect)
  $width = [Math]::Max(0, $rect.Right - $rect.Left)
  $height = [Math]::Max(0, $rect.Bottom - $rect.Top)

  $combinedText = (($textParts | Select-Object -Unique) -join " ").Trim()
  $looksLikeDialogClass = $className -in @("#32770", "bosa_sdm_XL9", "NUIDialog")
  $looksLikeOfficeDialog = $title -match '${SUPPORTED_DIALOG_TITLE_PATTERN}' -or $combinedText.Length -gt 0
  if (-not $looksLikeDialogClass -and -not $looksLikeOfficeDialog) { return $true }
  if (($className -eq "XLMAIN" -or $width -gt 1400 -or $height -gt 1000) -and $buttons.Count -le 1 -and $textParts.Count -eq 0) {
    return $true
  }

  $results.Add([ordered]@{
    handle = ([Int64]$hWnd).ToString()
    processId = [int]$windowProcessId
    processName = $processName
    title = $title
    className = $className
    text = $combinedText
    buttons = @($buttons | Select-Object -Unique)
    hasEdit = [bool]$hasEdit
    width = $width
    height = $height
  }) | Out-Null
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($results.Count -eq 0) {
  Write-Output "[]"
} else {
  $results | ConvertTo-Json -Depth 6
}
`;
  const result = await runPowerShell(script);
  if (!result.success) {
    return { dialogs: [], error: result.message || "弹窗探测脚本执行失败", rawOutput: result.output };
  }
  if (!result.output || !result.output.trim()) {
    return { dialogs: [], rawOutput: result.output };
  }
  try {
    const parsed = JSON.parse(result.output) as MacroDialogInfo[] | MacroDialogInfo;
    const items = Array.isArray(parsed) ? parsed : [parsed];
    return {
      dialogs: items.map((item) => ({ ...item, buttons: Array.isArray(item.buttons) ? item.buttons : [] })),
      rawOutput: result.output,
    };
  } catch (error) {
    return { dialogs: [], error: error instanceof Error ? error.message : String(error), rawOutput: result.output };
  }
}

async function invokeExcelDialogButton(handle: string, preferredButtons: string[]): Promise<boolean> {
  const preferred = preferredButtons.map((button) => `'${escapePowerShellSingleQuoted(button)}'`).join(", ");
  const script = `
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class VbeClickWin32 {
  public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumChildWindows(IntPtr hWnd, EnumChildProc callback, IntPtr lParam);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowTextLength(IntPtr hWnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);

  [DllImport("user32.dll")]
  public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindow(IntPtr hWnd);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
function Get-WindowTextSafe([IntPtr]$hWnd) {
  $len = [VbeClickWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][VbeClickWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][VbeClickWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
$target = [IntPtr]::new([long]'${escapePowerShellSingleQuoted(handle)}')
if (-not [VbeClickWin32]::IsWindow($target)) {
  Write-Output "true"
  exit 0
}
$buttons = New-Object System.Collections.Generic.List[object]
[VbeClickWin32]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [VbeClickWin32]::IsWindowVisible($child)) { return $true }
  $class = Get-ClassNameSafe $child
  if ($class -ne "Button") { return $true }
  $text = Get-WindowTextSafe $child
  if ([string]::IsNullOrWhiteSpace($text)) { return $true }
  $buttons.Add(@{ Handle = $child; Text = $text }) | Out-Null
  return $true
}, [IntPtr]::Zero) | Out-Null

$preferred = @(${preferred})
$match = $null
foreach ($p in $preferred) {
  $match = $buttons | Where-Object { $_.Text -eq $p } | Select-Object -First 1
  if ($null -ne $match) { break }
}
if ($null -eq $match) {
  # 中文 Excel 按钮文本带热键后缀（如"结束(&E)"），去掉 & 后做包含匹配
  foreach ($p in $preferred) {
    $pNorm = ($p -replace '&', '')
    $match = $buttons | Where-Object { (($_.Text -replace '&', '') -like "*$pNorm*") } | Select-Object -First 1
    if ($null -ne $match) { break }
  }
}
if ($null -eq $match -and $buttons.Count -eq 1) {
  $match = $buttons[0]
}
if ($null -eq $match) {
  Write-Output "false"
  exit 0
}
try {
  [void][VbeClickWin32]::SendMessage($match.Handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
  Start-Sleep -Milliseconds 200
  if (-not [VbeClickWin32]::IsWindow($target) -or -not [VbeClickWin32]::IsWindowVisible($target)) {
    Write-Output "true"
    exit 0
  }
  try {
    $element = [System.Windows.Automation.AutomationElement]::FromHandle($match.Handle)
    if ($null -ne $element) {
      $invoke = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
      $invoke.Invoke()
    }
  } catch {}
  Start-Sleep -Milliseconds 200
  # 不做键盘模拟（SendKeys 会落入用户前台窗口），消息级点击 + UIA 已覆盖全部实测场景
  if (-not [VbeClickWin32]::IsWindow($target) -or -not [VbeClickWin32]::IsWindowVisible($target)) {
    Write-Output "true"
  } else {
    Write-Output "false"
  }
} catch {
  Write-Output "false"
}
`;
  const result = await runPowerShell(script);
  return result.success && result.output?.trim() === "true";
}

/**
 * 编译错误弹窗点击"确定"后 VBE 会进入中断模式（Run 调用挂起不返回）。
 * 通过新鲜 COM 附着程序化执行 VBE 菜单：中断（id=189，运行态时先转中断）→
 * 重新设置（id=228，中断态专属项）。全程 CommandBars/Win32 消息完成，
 * 不使用任何键盘模拟，不影响用户前台操作。
 * 重置确认框（"该操作将重新设置工程"）用 BM_CLICK 自动点"确定"。
 */
export async function resetVbeInterrupt(processId: number): Promise<boolean> {
  for (let attempt = 0; attempt < 3; attempt++) {
    const out = await resetVbeViaCommandBars(processId);
    if (out === "novbe" || out === "reset") return out === "reset" || attempt > 0;
    await sleep(1000);
  }
  return false;
}

async function resetVbeViaCommandBars(processId: number): Promise<string> {
  const script = `
$ErrorActionPreference = "Stop"
${VBE_RESET_PS}
$targetPid = ${processId}
$state = Get-VbeState $targetPid
Write-Output "STATE=$state"
if ($state -eq "novbe" -or $state -eq "design") { Write-Output "nobreak"; exit 0 }
$excel = Attach-Excel $targetPid
$vbe = $excel.VBE
# 运行态先中断，再重置
if ($state -eq "running") {
  $brk = $vbe.CommandBars.FindControl([Type]::Missing, 189)
  if ($null -ne $brk) { $brk.Execute(); Start-Sleep -Milliseconds 800 }
}
$reset = $vbe.CommandBars.FindControl([Type]::Missing, 228)
if ($null -eq $reset) { Write-Output "noreset"; exit 0 }
$reset.Execute()
Write-Output "EXECUTED"
Start-Sleep -Milliseconds 600
# 自动点击"该操作将重新设置工程"确认框（BM_CLICK，无需键盘/前台）
for ($i = 0; $i -lt 5; $i++) {
  $clicked = Invoke-ResetConfirmClick $targetPid
  if ($clicked) { break }
  Start-Sleep -Milliseconds 500
}
Start-Sleep -Milliseconds 800
$final = Get-VbeState $targetPid
Write-Output "FINAL=$final"
if ($final -eq "design" -or $final -eq "novbe") { Write-Output "reset" } else { Write-Output "stillbreak" }
`;
  const result = await runPowerShell(script, 45000);
  const lines = (result.output || "").trim().split("\n");
  return lines[lines.length - 1] || "";
}

const VBE_RESET_PS = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class VbeResetWin32 {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
  [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc callback, IntPtr lParam);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}
"@
function Attach-Excel([int]$procId) {
  $script:foundHwnd = [IntPtr]::Zero
  [VbeResetWin32]::EnumWindows({
    param($h, $l)
    if (-not [VbeResetWin32]::IsWindowVisible($h)) { return $true }
    $p = [uint32]0; [void][VbeResetWin32]::GetWindowThreadProcessId($h, [ref]$p)
    if ($p -ne $procId) { return $true }
    $sb = New-Object System.Text.StringBuilder 256
    [void][VbeResetWin32]::GetClassName($h, $sb, 256)
    if ($sb.ToString() -eq "XLMAIN") { $script:foundHwnd = $h; return $false }
    return $true
  }, [IntPtr]::Zero) | Out-Null
  $script:omHwnd = [IntPtr]::Zero
  [VbeResetWin32]::EnumChildWindows($script:foundHwnd, {
    param($h, $l)
    $sb = New-Object System.Text.StringBuilder 256
    [void][VbeResetWin32]::GetClassName($h, $sb, 256)
    if ($sb.ToString() -eq "EXCEL7") { $script:omHwnd = $h; return $false }
    return $true
  }, [IntPtr]::Zero) | Out-Null
  $t = if ($script:omHwnd -ne [IntPtr]::Zero) { $script:omHwnd } else { $script:foundHwnd }
  $guid = [Guid]::Parse("00020400-0000-0000-C000-000000000046")
  $obj = $null
  $hr = [VbeResetWin32]::AccessibleObjectFromWindow($t, [UInt32]4294967280, [ref]$guid, [ref]$obj)
  if ($hr -ne 0) { throw "attach fail" }
  return $obj.Application
}
function Get-VbeTitle([int]$procId) {
  $script:vt = ""
  [VbeResetWin32]::EnumWindows({
    param($h, $l)
    if (-not [VbeResetWin32]::IsWindowVisible($h)) { return $true }
    $p = [uint32]0; [void][VbeResetWin32]::GetWindowThreadProcessId($h, [ref]$p)
    if ($p -ne $procId) { return $true }
    $sb = New-Object System.Text.StringBuilder 512
    [void][VbeResetWin32]::GetWindowText($h, $sb, 512)
    $t = $sb.ToString()
    if ($t -like "Microsoft Visual Basic*") { $script:vt = $t; return $false }
    return $true
  }, [IntPtr]::Zero) | Out-Null
  return $script:vt
}
function Get-VbeState([int]$procId) {
  $title = Get-VbeTitle $procId
  if ($title -eq "") { return "novbe" }
  if ($title -match '\[中断\]|\[break\]') { return "break" }
  if ($title -match '\[正在运行\]|\[running\]') { return "running" }
  return "design"
}
function Invoke-ResetConfirmClick([int]$procId) {
  $script:confirmHwnd = [IntPtr]::Zero
  [VbeResetWin32]::EnumWindows({
    param($h, $l)
    if (-not [VbeResetWin32]::IsWindowVisible($h)) { return $true }
    $p = [uint32]0; [void][VbeResetWin32]::GetWindowThreadProcessId($h, [ref]$p)
    if ($p -ne $procId) { return $true }
    $cls = New-Object System.Text.StringBuilder 256
    [void][VbeResetWin32]::GetClassName($h, $cls, 256)
    if ($cls.ToString() -ne "#32770") { return $true }
    $tb = New-Object System.Text.StringBuilder 512
    [void][VbeResetWin32]::GetWindowText($h, $tb, 512)
    if ($tb.ToString() -match '^Microsoft Visual Basic( for Applications)?$') {
      $txb = New-Object System.Text.StringBuilder 1024
      $script:dialogText = ""
      [VbeResetWin32]::EnumChildWindows($h, {
        param($c, $l2)
        $ccb = New-Object System.Text.StringBuilder 256
        [void][VbeResetWin32]::GetClassName($c, $ccb, 256)
        if ($ccb.ToString() -eq "Static") {
          $stb = New-Object System.Text.StringBuilder 1024
          [void][VbeResetWin32]::GetWindowText($c, $stb, 1024)
          $script:dialogText += $stb.ToString()
        }
        return $true
      }, [IntPtr]::Zero) | Out-Null
      if ($script:dialogText -match '重新设置工程|reset the project') { $script:confirmHwnd = $h; return $false }
    }
    return $true
  }, [IntPtr]::Zero) | Out-Null
  if ($script:confirmHwnd -eq [IntPtr]::Zero) { return $false }
  $script:okBtn = [IntPtr]::Zero
  [VbeResetWin32]::EnumChildWindows($script:confirmHwnd, {
    param($c, $l2)
    $ccb = New-Object System.Text.StringBuilder 256
    [void][VbeResetWin32]::GetClassName($c, $ccb, 256)
    if ($ccb.ToString() -eq "Button") {
      $stb = New-Object System.Text.StringBuilder 256
      [void][VbeResetWin32]::GetWindowText($c, $stb, 256)
      $tn = $stb.ToString().Replace("&", "")
      if ($tn -match '^确定$|^OK$|^是$|^Yes$') { $script:okBtn = $c; return $false }
    }
    return $true
  }, [IntPtr]::Zero) | Out-Null
  if ($script:okBtn -eq [IntPtr]::Zero) { return $false }
  [void][VbeResetWin32]::SendMessage($script:okBtn, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
  return $true
}
`;



function enrichDialogInfo(dialog: MacroDialogInfo): MacroDialogInfo & {
  summary: string;
  recommendedAction: string | null;
  recommendedButtons: string[];
} {
  const normalizedDialog: MacroDialogInfo = { ...dialog, kind: dialog.kind || classifyMacroDialog(dialog) };
  const recommended = pickDialogAction(normalizedDialog);
  return {
    ...normalizedDialog,
    summary: summarizeDialog(normalizedDialog),
    recommendedAction: recommended?.actionLabel || null,
    recommendedButtons: recommended?.buttons || [],
  };
}

/**
 * 依据弹窗分类与交互策略解析自动处理动作。
 * 错误弹窗始终自动处理；info/确认/InputBox 仅在 policy.mode === "auto" 时处理。
 */
function resolveAutoAction(
  dialog: MacroDialogInfo,
  policy?: MacroDialogPolicy
): { buttons: string[]; actionLabel: string; terminal: boolean } | { fill: true } | null {
  const kind = dialog.kind || classifyMacroDialog(dialog);
  if (kind === "vb_runtime_error" || kind === "error") {
    return pickDialogAction(dialog);
  }
  if (!policy || policy.mode !== "auto") return null;
  // info/确认/InputBox 仅对真正的模态对话框自动处理，避免误点 VBE 主窗口等复杂界面
  if (!["#32770", "bosa_sdm_XL9", "NUIDialog", "bosa_sdm_msword"].includes(dialog.className)) return null;
  // InputBox：有输入框且提供了期望输入 → 填入并提交（优先于确认框处理）
  if (dialog.hasEdit && policy.inputValue) return { fill: true };
  if (kind === "info") {
    return { buttons: ["确定", "OK", "继续", "Continue"], actionLabel: "确定", terminal: false };
  }
  if (kind === "confirmation") {
    return resolveDialogAction(dialog, undefined, policy.confirmButton || "取消");
  }
  return null;
}

export async function listExcelDialogs(): Promise<ExcelComResult> {
  const inspection = await inspectExcelDialogs();
  if (inspection.error && inspection.dialogs.length === 0) {
    return { success: false, message: `弹窗探测失败：${inspection.error}`, details: { rawOutput: inspection.rawOutput || null } };
  }
  const dialogs = inspection.dialogs.map((dialog) => enrichDialogInfo(dialog));
  return {
    success: true,
    message: dialogs.length > 0 ? `检测到 ${dialogs.length} 个 Excel/VBA 弹窗` : "当前没有检测到 Excel/VBA 弹窗",
    output: JSON.stringify(dialogs, null, 2),
    details: { count: dialogs.length, dialogs, inspectionError: inspection.error || null },
  };
}

export async function clickExcelDialog(handle: string, action?: string, buttonText?: string): Promise<ExcelComResult> {
  const inspection = await inspectExcelDialogs();
  if (inspection.error && inspection.dialogs.length === 0) {
    return { success: false, message: `弹窗探测失败：${inspection.error}`, details: { rawOutput: inspection.rawOutput || null } };
  }
  const target = inspection.dialogs.find((dialog) => dialog.handle === handle);
  if (!target) {
    return { success: false, message: `未找到句柄为 ${handle} 的弹窗。请先调用 excel_list_dialogs 刷新当前弹窗列表。` };
  }
  const dialog = enrichDialogInfo(target);
  const resolvedAction = resolveDialogAction(dialog, action, buttonText);
  if (!resolvedAction) {
    return { success: false, message: `无法为当前弹窗确定处理动作。可用按钮：${dialog.buttons.join(" / ") || "无"}`, details: { dialog } };
  }
  const handled = await invokeExcelDialogButton(dialog.handle, resolvedAction.buttons);
  if (!handled) {
    return {
      success: false,
      message: `未能点击弹窗按钮「${resolvedAction.actionLabel}」。可用按钮：${dialog.buttons.join(" / ") || "无"}`,
      details: { dialog, requestedAction: action || null, requestedButtonText: buttonText || null, preferredButtons: resolvedAction.buttons },
    };
  }
  // 编译错误弹窗点"确定"后 VBE 进入中断模式，Run 会一直挂起，需要重置
  if (dialog.kind === "vb_runtime_error" && /编译错误|compile error/i.test(`${dialog.title} ${dialog.text}`)) {
    await resetVbeInterrupt(dialog.processId || 0);
  }
  return {
    success: true,
    message: `已点击弹窗按钮「${resolvedAction.actionLabel}」`,
    details: { dialog: { ...dialog, autoHandled: true, autoAction: resolvedAction.actionLabel }, requestedAction: action || null, requestedButtonText: buttonText || null, preferredButtons: resolvedAction.buttons },
  };
}

/** 向弹窗输入框写入文本。找到弹窗中第一个可见 Edit 控件，写入文本并可选提交 */
export async function fillDialogInput(handle: string, text: string, submit = false): Promise<ExcelComResult> {
  const inspection = await inspectExcelDialogs();
  if (inspection.error && inspection.dialogs.length === 0) {
    return { success: false, message: `弹窗探测失败：${inspection.error}`, details: { rawOutput: inspection.rawOutput || null } };
  }
  const target = inspection.dialogs.find((dialog) => dialog.handle === handle);
  if (!target) {
    return { success: false, message: `未找到句柄为 ${handle} 的弹窗。请先调用 excel_list_dialogs 刷新当前弹窗列表。` };
  }

  const escapedText = escapePowerShellSingleQuoted(text);
  const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class VbeFillWin32 {
  public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumChildWindows(IntPtr hWnd, EnumChildProc callback, IntPtr lParam);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowTextLength(IntPtr hWnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int maxCount);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetForegroundWindow(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern IntPtr GetForegroundWindow();

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

  [DllImport("user32.dll", EntryPoint = "SendMessage")]
  public static extern IntPtr SendMessageInt(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

  [DllImport("user32.dll")]
  public static extern bool SetFocus(IntPtr hWnd);
}
"@
function Get-WindowTextSafe([IntPtr]$hWnd) {
  $len = [VbeFillWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][VbeFillWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][VbeFillWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
$target = [IntPtr]::new([long]'${escapePowerShellSingleQuoted(handle)}')
if (-not [VbeFillWin32]::IsWindowVisible($target)) {
  Write-Output (@{ success = $false; error = "弹窗不可见" } | ConvertTo-Json -Compress)
  exit 0
}
$script:edit = $null
[VbeFillWin32]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [VbeFillWin32]::IsWindowVisible($child)) { return $true }
  $class = Get-ClassNameSafe $child
  if ($class -in @("Edit", "RichEdit20W", "RichEdit50W", "RICHEDIT50W", "msctls_hotkey32")) {
    if ($null -eq $script:edit) { $script:edit = $child }
    return $false
  }
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($null -eq $script:edit) {
  Write-Output (@{ success = $false; error = "未在弹窗中找到输入框控件" } | ConvertTo-Json -Compress)
  exit 0
}
# 统一改用 $script:edit 引用（脚本块委托内赋值不会传回本地作用域）
$edit = $script:edit
[void][VbeFillWin32]::SetForegroundWindow($target)
Start-Sleep -Milliseconds 100
[void][VbeFillWin32]::SetFocus($edit)
# 先尝试直接设置文本
$set = [VbeFillWin32]::SendMessage($edit, 0x000C, [IntPtr]::Zero, '${escapedText}')
Start-Sleep -Milliseconds 100
# 提交：优先对"确定/OK"按钮发 BM_CLICK（不依赖键盘焦点），失败再用 SendKeys
$okBtn = $null
[VbeFillWin32]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [VbeFillWin32]::IsWindowVisible($child)) { return $true }
  if ((Get-ClassNameSafe $child) -ne "Button") { return $true }
  $t = (Get-WindowTextSafe $child) -replace '&', ''
  if ($t -like '*确定*' -or $t -like '*OK*') { $script:okBtn = $child; return $false }
  return $true
}, [IntPtr]::Zero) | Out-Null
$clicked = $false
if ($script:okBtn -ne $null) {
  [void][VbeFillWin32]::SendMessageInt($script:okBtn, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
  Start-Sleep -Milliseconds 250
  $clicked = -not [VbeFillWin32]::IsWindowVisible($target)
}
if (-not $clicked -and ${submit ? "$true" : "$false"}) {
  # 键盘模拟仅当目标弹窗已经是前台窗口时才允许，绝不波及用户正在操作的其他窗口
  $fg = [VbeFillWin32]::GetForegroundWindow()
  if ($fg -eq $target) {
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.SendKeys('^a${escapedText}')
    Start-Sleep -Milliseconds 100
    $wshell.SendKeys('{ENTER}')
    Start-Sleep -Milliseconds 200
  }
}
$actual = Get-WindowTextSafe $edit
Write-Output (@{ success = $true; editText = $actual; clickedOk = $clicked } | ConvertTo-Json -Compress)
`;
  const result = await runPowerShell(script);
  if (!result.success) {
    return { success: false, message: `输入文本失败：${result.message}`, output: result.output };
  }
  try {
    const parsed = JSON.parse(result.output || "{}") as { success: boolean; error?: string; editText?: string };
    if (!parsed.success) {
      return { success: false, message: parsed.error || "输入文本失败" };
    }
    return {
      success: true,
      message: submit ? `已输入文本并提交：${text}` : `已输入文本：${text}`,
      details: { editText: parsed.editText },
    };
  } catch {
    return { success: true, message: `已尝试输入文本：${text}`, output: result.output };
  }
}

export async function runMacroWithDialogHandling(
  filePath: string,
  macroName: string,
  options: { timeoutMs?: number; captureResultRange?: string; excelProcessId?: number; dialogPolicy?: MacroDialogPolicy } = {}
): Promise<ExcelComResult> {
  const normalizedMacroName = macroName.trim();
  if (!normalizedMacroName) {
    return { success: false, message: "macroName 不能为空。" };
  }

  const wbName = basename(filePath);
  const preCheck = await ensureExcelRunning(wbName, options.excelProcessId);
  if (preCheck) return preCheck;

  const dialogPolicy = options.dialogPolicy;

  const timeoutMs = options.timeoutMs ?? 45000;
  const statusPath = join(tmpdir(), `vba_macro_status_${randomBytes(8).toString("hex")}.json`);
  const runnerPath = join(tmpdir(), `vba_macro_runner_${randomBytes(8).toString("hex")}.ps1`);
  const bom = "\uFEFF";
  const utf8Header = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
`;
  await writeFile(
    runnerPath,
    bom + utf8Header + buildMacroRunnerScript({ filePath, workbookName: wbName, macroName: normalizedMacroName, statusPath, excelProcessId: options.excelProcessId }),
    "utf-8"
  );

  const child = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", runnerPath], {
    windowsHide: true,
    stdio: "ignore",
  });

  const collectedDialogs: MacroDialogInfo[] = [];
  const dialogState = new Map<string, { fingerprint: string; firstSeen: number; handled: boolean; recorded: boolean }>();
  const recordedFingerprints = new Set<string>();
  const startedAt = Date.now();
  let lastDialogInspectionError: string | undefined;

  try {
    while (Date.now() - startedAt < timeoutMs) {
      const status = await readJsonFile<MacroRunnerStatus>(statusPath);
      if (status) {
        return buildMacroRunResult(status, normalizedMacroName, wbName, startedAt, collectedDialogs, lastDialogInspectionError, options.captureResultRange);
      }

      const inspection = await inspectExcelDialogs();
      if (inspection.error) {
        lastDialogInspectionError = inspection.error;
      }
      for (const rawDialog of inspection.dialogs) {
        const dialog: MacroDialogInfo = { ...rawDialog, kind: classifyMacroDialog(rawDialog) };
        const key = dialog.handle;
        const fingerprint = buildDialogFingerprint(dialog);
        const existing = dialogState.get(key);
        if (!existing || existing.fingerprint !== fingerprint) {
          dialogState.set(key, { fingerprint, firstSeen: Date.now(), handled: false, recorded: false });
          continue;
        }
        if (existing.handled || Date.now() - existing.firstSeen < MACRO_DIALOG_STABLE_MS) {
          continue;
        }
        if (!existing.recorded) {
          if (!recordedFingerprints.has(fingerprint)) {
            collectedDialogs.push({ ...dialog });
            recordedFingerprints.add(fingerprint);
          }
          existing.recorded = true;
        }

        const action = resolveAutoAction(dialog, dialogPolicy);
        if (!action) continue;

        if ("fill" in action) {
          // InputBox：填入期望输入并提交
          const filled = await fillDialogInput(dialog.handle, dialogPolicy!.inputValue!, true);
          if (!filled.success) continue;
          dialog.autoHandled = true;
          dialog.autoAction = `输入文本: ${dialogPolicy!.inputValue}`;
          const fillIndex = collectedDialogs.findIndex((item) => buildDialogFingerprint(item) === fingerprint);
          if (fillIndex >= 0) {
            collectedDialogs[fillIndex] = { ...collectedDialogs[fillIndex], autoHandled: true, autoAction: dialog.autoAction };
          } else {
            collectedDialogs.push(dialog);
            recordedFingerprints.add(fingerprint);
          }
          existing.handled = true;
          continue;
        }

        const handled = await invokeExcelDialogButton(dialog.handle, action.buttons);
        if (!handled) continue;

        dialog.autoHandled = true;
        dialog.autoAction = action.actionLabel;
        const dialogIndex = collectedDialogs.findIndex((item) => buildDialogFingerprint(item) === fingerprint);
        if (dialogIndex >= 0) {
          collectedDialogs[dialogIndex] = { ...collectedDialogs[dialogIndex], autoHandled: true, autoAction: action.actionLabel };
        } else {
          collectedDialogs.push(dialog);
          recordedFingerprints.add(fingerprint);
        }
        existing.handled = true;

        // 编译错误弹窗点"确定"后 VBE 进入中断模式，Run 会一直挂起，需要重置
        if (dialog.kind === "vb_runtime_error" && /编译错误|compile error/i.test(`${dialog.title} ${dialog.text}`)) {
          await resetVbeInterrupt(dialog.processId || 0);
        }

        if (action.terminal) {
          await sleep(900);
        }
      }

      if (child.exitCode !== null) {
        break;
      }
      await sleep(MACRO_DIALOG_POLL_MS);
    }

    if (child.exitCode !== null) {
      const finalStatus = await waitForRunnerStatus(statusPath, MACRO_RUNNER_STATUS_GRACE_MS);
      if (finalStatus) {
        return buildMacroRunResult(finalStatus, normalizedMacroName, wbName, startedAt, collectedDialogs, lastDialogInspectionError, options.captureResultRange);
      }
      const finishedDetails = {
        macroName: normalizedMacroName,
        workbook: wbName,
        durationMs: Date.now() - startedAt,
        dialogs: collectedDialogs,
        dialogInspectionError: lastDialogInspectionError,
        exitCode: child.exitCode,
        statusMissing: true,
      };
      if (child.exitCode === 0) {
        const autoHandledCount = collectedDialogs.filter((dialog) => dialog.autoHandled).length;
        return {
          success: true,
          message: autoHandledCount > 0
            ? `宏 ${normalizedMacroName} 执行完成（已自动处理 ${autoHandledCount} 个弹窗）`
            : `宏 ${normalizedMacroName} 执行完成`,
          details: finishedDetails,
        };
      }
      return { success: false, message: `宏执行进程已结束，但未返回可解析的状态（exit code: ${child.exitCode}）。`, details: finishedDetails };
    }

    const timeoutDetails = {
      macroName: normalizedMacroName,
      workbook: wbName,
      durationMs: Date.now() - startedAt,
      dialogs: collectedDialogs,
      dialogInspectionError: lastDialogInspectionError,
    };
    const timeoutMessage = lastDialogInspectionError
      ? `宏执行期间弹窗探测失败：${lastDialogInspectionError}`
      : collectedDialogs.length > 0
        ? `宏执行超时，期间检测到并处理了弹窗。最后一个弹窗：${summarizeDialog(collectedDialogs[collectedDialogs.length - 1])}`
        : "宏执行超时，Excel 可能仍在运行宏或弹出了未能自动处理的对话框。";
    return { success: false, message: timeoutMessage, details: timeoutDetails };
  } finally {
    if (child.exitCode === null) {
      child.kill();
    }
    unlink(runnerPath).catch(() => {});
    unlink(statusPath).catch(() => {});
  }
}

function buildMacroRunResult(
  status: MacroRunnerStatus,
  macroName: string,
  workbook: string,
  startedAt: number,
  dialogs: MacroDialogInfo[],
  dialogInspectionError?: string,
  captureResultRange?: string
): ExcelComResult {
  const durationMs = Date.now() - startedAt;
  const details: Record<string, unknown> = {
    ...(status.details || {}),
    macroName,
    workbook,
    durationMs,
    dialogs,
    dialogInspectionError,
  };

  // 如果指定了结果捕获区域，尝试读取
  if (captureResultRange) {
    // 这里不直接读取，因为读取需要在主 PowerShell 中执行；留给调用方或后续实现
    details.captureResultRange = captureResultRange;
  }

  if (status.state === "success") {
    const autoHandledCount = dialogs.filter((dialog) => dialog.autoHandled).length;
    return {
      success: true,
      message: autoHandledCount > 0
        ? `${status.message}（已自动处理 ${autoHandledCount} 个弹窗）`
        : status.message,
      output: status.details?.runFormat as string | undefined,
      details,
    };
  }
  const runtimeDialog = [...dialogs].reverse().find((dialog) => dialog.kind === "vb_runtime_error" || dialog.kind === "error");
  const failurePrefix = runtimeDialog
    ? `检测到弹窗：${summarizeDialog(runtimeDialog)}`
    : friendlyError(status.message);
  const hresult = typeof status.details?.hresult === "string" && status.details.hresult
    ? `（HRESULT: ${status.details.hresult}）`
    : "";
  return {
    success: false,
    message: `${failurePrefix}${hresult}`,
    output: typeof status.details?.rawError === "string" ? status.details.rawError : status.message,
    details,
  };
}

