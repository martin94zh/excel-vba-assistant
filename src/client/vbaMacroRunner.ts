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

import { escapePowerShellSingleQuoted, friendlyError, runPowerShell } from "../runtime/powershell";
import type { ExcelComResult } from "../runtime/powershell";
import { ensureExcelRunning } from "../runtime/powershell";

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
  hasInputField?: boolean;
  kind?: "input" | "vb_runtime_error" | "error" | "confirmation" | "info" | "unknown";
  autoHandled?: boolean;
  autoAction?: string;
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
const SUPPORTED_DIALOG_TITLE_PATTERN = "(?i)(microsoft excel|microsoft visual basic|excel|visual basic|wps|kingsoft|et)";

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
  // InputBox（有可编辑文本输入框）优先识别，避免被误判为 confirmation 而自动取消
  if (dialog.hasInputField) {
    return "input";
  }
  const haystack = `${dialog.title}\n${dialog.text}\n${dialog.buttons.join(" ")}`.toLowerCase();
  if (
    haystack.includes("microsoft visual basic")
    || haystack.includes("运行时错误")
    || haystack.includes("compile error")
    || haystack.includes("编译错误")
    || haystack.includes("debug")
  ) {
    return "vb_runtime_error";
  }
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
  const kind = dialog.kind || classifyMacroDialog(dialog);
  switch (kind) {
    case "input":
      // InputBox 需要用户输入，不自动点击取消，避免宏直接 Exit Sub
      // 由 autoFillInputs 参数或外部 excel_fill_dialog/excel_click_dialog 处理
      return null;
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
    case "debug":
    case "调试":
      return { buttons: ["调试", "Debug"], actionLabel: "调试", terminal: false };
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
  $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
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

public static class JrVbeWin32 {
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
  $len = [JrVbeWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][JrVbeWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrVbeWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ProcessNameSafe([int]$targetProcessId) {
  try { return (Get-Process -Id $targetProcessId -ErrorAction Stop).ProcessName } catch { return "" }
}
$results = New-Object System.Collections.Generic.List[object]
[JrVbeWin32]::EnumWindows({
  param($hWnd, $lParam)
  if (-not [JrVbeWin32]::IsWindowVisible($hWnd)) { return $true }
  $windowProcessId = 0; [void][JrVbeWin32]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
  $title = Get-WindowTextSafe $hWnd
  $className = Get-ClassNameSafe $hWnd
  $processName = Get-ProcessNameSafe $windowProcessId
  if ($processName -notmatch '${SUPPORTED_DIALOG_PROCESS_PATTERN}') { return $true }
  $buttons = New-Object System.Collections.Generic.List[string]
  $textParts = New-Object System.Collections.Generic.List[string]
  $hasInputField = $false
  [JrVbeWin32]::EnumChildWindows($hWnd, {
    param($childHwnd, $childLparam)
    if (-not [JrVbeWin32]::IsWindowVisible($childHwnd)) { return $true }
    $childClass = Get-ClassNameSafe $childHwnd
    $childText = Get-WindowTextSafe $childHwnd
    if ($childClass -eq "Button") {
      if (-not [string]::IsNullOrWhiteSpace($childText) -and -not $buttons.Contains($childText)) { $buttons.Add($childText) | Out-Null }
      return $true
    }
    if ($childClass -in @("Edit", "RichEdit20W", "RichEdit50W", "RICHEDIT50W")) {
      $script:hasInputField = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($childText) -and -not $textParts.Contains($childText)) {
      $textParts.Add($childText) | Out-Null
    }
    return $true
  }, [IntPtr]::Zero) | Out-Null

  if ($buttons.Count -eq 0) { return $true }

  $rect = New-Object JrVbeWin32+RECT
  [void][JrVbeWin32]::GetWindowRect($hWnd, [ref]$rect)
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
    width = $width
    height = $height
    hasInputField = $hasInputField
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

public static class JrVbeClickWin32 {
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
  $len = [JrVbeClickWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][JrVbeClickWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrVbeClickWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
$target = [IntPtr]::new([long]'${escapePowerShellSingleQuoted(handle)}')
if (-not [JrVbeClickWin32]::IsWindow($target)) {
  Write-Output "true"
  exit 0
}
$buttons = New-Object System.Collections.Generic.List[object]
[JrVbeClickWin32]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [JrVbeClickWin32]::IsWindowVisible($child)) { return $true }
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
if ($null -eq $match -and $buttons.Count -eq 1) {
  $match = $buttons[0]
}
if ($null -eq $match) {
  Write-Output "false"
  exit 0
}
try {
  [void][JrVbeClickWin32]::SendMessage($match.Handle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
  Start-Sleep -Milliseconds 150
  if (-not [JrVbeClickWin32]::IsWindow($target) -or -not [JrVbeClickWin32]::IsWindowVisible($target)) {
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
  Start-Sleep -Milliseconds 150
  if (-not [JrVbeClickWin32]::IsWindow($target) -or -not [JrVbeClickWin32]::IsWindowVisible($target)) {
    Write-Output "true"
    exit 0
  }
  try {
    [void][JrVbeClickWin32]::SetForegroundWindow($target)
    $wshell = New-Object -ComObject WScript.Shell
    [void]$wshell.AppActivate((Get-WindowTextSafe $target))
    $wshell.SendKeys("{ENTER}")
  } catch {}
  Start-Sleep -Milliseconds 150
  if (-not [JrVbeClickWin32]::IsWindow($target) -or -not [JrVbeClickWin32]::IsWindowVisible($target)) {
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

public static class JrVbeFillWin32 {
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

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

  [DllImport("user32.dll")]
  public static extern bool SetFocus(IntPtr hWnd);
}
"@
function Get-WindowTextSafe([IntPtr]$hWnd) {
  $len = [JrVbeFillWin32]::GetWindowTextLength($hWnd)
  $sb = New-Object System.Text.StringBuilder ([Math]::Max($len + 1, 8))
  [void][JrVbeFillWin32]::GetWindowText($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
function Get-ClassNameSafe([IntPtr]$hWnd) {
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrVbeFillWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
  return $sb.ToString().Trim()
}
$target = [IntPtr]::new([long]'${escapePowerShellSingleQuoted(handle)}')
if (-not [JrVbeFillWin32]::IsWindowVisible($target)) {
  Write-Output (@{ success = $false; error = "弹窗不可见" } | ConvertTo-Json -Compress)
  exit 0
}
$edit = $null
[JrVbeFillWin32]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [JrVbeFillWin32]::IsWindowVisible($child)) { return $true }
  $class = Get-ClassNameSafe $child
  if ($class -in @("Edit", "RichEdit20W", "RichEdit50W", "RICHEDIT50W", "msctls_hotkey32")) {
    if ($null -eq $edit) { $edit = $child }
    return $false
  }
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($null -eq $edit) {
  Write-Output (@{ success = $false; error = "未在弹窗中找到输入框控件" } | ConvertTo-Json -Compress)
  exit 0
}
[void][JrVbeFillWin32]::SetForegroundWindow($target)
Start-Sleep -Milliseconds 100
[void][JrVbeFillWin32]::SetFocus($edit)
# 先尝试直接设置文本
$set = [JrVbeFillWin32]::SendMessage($edit, 0x000C, [IntPtr]::Zero, '${escapedText}')
Start-Sleep -Milliseconds 100
# 如果 WM_SETTEXT 失败或需要提交，再用 SendKeys
if ($set -eq [IntPtr]::Zero -or ${submit ? "$true" : "$false"}) {
  $wshell = New-Object -ComObject WScript.Shell
  [void]$wshell.AppActivate((Get-WindowTextSafe $target))
  $wshell.SendKeys('^a${escapedText}')
  Start-Sleep -Milliseconds 100
}
if (${submit ? "$true" : "$false"}) {
  $wshell = New-Object -ComObject WScript.Shell
  $wshell.SendKeys('{ENTER}')
  Start-Sleep -Milliseconds 200
}
$actual = Get-WindowTextSafe $edit
Write-Output (@{ success = $true; editText = $actual } | ConvertTo-Json -Compress)
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

/**
 * 向 InputBox 的编辑框写入文本并点击确定。
 * 用于宏执行期间自动填充 InputBox 弹窗。
 */
async function fillInputBoxAndSubmit(handle: string, text: string): Promise<boolean> {
  const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class JrInputBoxFill {
  public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool EnumChildWindows(IntPtr hWnd, EnumChildProc callback, IntPtr lParam);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);
  [DllImport("user32.dll")]
  public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
$target = [IntPtr]::new([long]'${escapePowerShellSingleQuoted(handle)}')
$editHandle = [IntPtr]::Zero
$okHandle = [IntPtr]::Zero
[JrInputBoxFill]::EnumChildWindows($target, {
  param($child, $lp)
  if (-not [JrInputBoxFill]::IsWindowVisible($child)) { return $true }
  $sb = New-Object System.Text.StringBuilder 256
  [void][JrInputBoxFill]::GetClassName($child, $sb, 256)
  $class = $sb.ToString()
  if ($class -in @("Edit", "RichEdit20W", "RichEdit50W", "RICHEDIT50W")) {
    $script:editHandle = $child
  }
  $sb2 = New-Object System.Text.StringBuilder 256
  [void][JrInputBoxFill]::GetWindowText($child, $sb2, 256)
  $btnText = $sb2.ToString()
  if ($class -eq "Button" -and ($btnText -eq "确定" -or $btnText -eq "OK")) {
    $script:okHandle = $child
  }
  return $true
}, [IntPtr]::Zero) | Out-Null
if ($editHandle -eq [IntPtr]::Zero) { Write-Output "false"; exit 0 }
# WM_SETTEXT = 0x000C
[void][JrInputBoxFill]::SendMessage($editHandle, 0x000C, [IntPtr]::Zero, '${escapePowerShellSingleQuoted(text)}')
Start-Sleep -Milliseconds 200
# 点击确定按钮
if ($okHandle -ne [IntPtr]::Zero) {
  # BM_CLICK = 0x00F5
  [void][JrInputBoxFill]::SendMessage($okHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
} else {
  # 回退：按回车键
  [void][JrInputBoxFill]::SetForegroundWindow($target)
  $wshell = New-Object -ComObject WScript.Shell
  [void]$wshell.AppActivate("InputBox")
  $wshell.SendKeys("{ENTER}")
}
Start-Sleep -Milliseconds 300
Write-Output "true"
`;
  const result = await runPowerShell(script, 8000);
  return result.success && result.output?.trim() === "true";
}

// ============================================================
// 宏运行 Session 管理（支持交互式弹窗处理）
// ============================================================

interface MacroSession {
  sessionId: string;
  child: import("child_process").ChildProcess;
  statusPath: string;
  runnerPath: string;
  startedAt: number;
  deadline: number;
  createdAt: number;
  macroName: string;
  workbookName: string;
  collectedDialogs: MacroDialogInfo[];
  dialogState: Map<string, { fingerprint: string; firstSeen: number; handled: boolean; recorded: boolean }>;
  recordedFingerprints: Set<string>;
  lastDialogInspectionError?: string;
  captureResultRange?: string;
  autoFillInputs?: string[];
}

const macroSessions = new Map<string, MacroSession>();
const MACRO_SESSION_TTL_MS = 30 * 60 * 1000;

function cleanupSession(session: MacroSession): void {
  if (session.child.exitCode === null) {
    session.child.kill();
  }
  unlink(session.runnerPath).catch(() => {});
  unlink(session.statusPath).catch(() => {});
}

function cleanupExpiredSessions(): void {
  const now = Date.now();
  for (const [id, session] of macroSessions) {
    if (now - session.createdAt > MACRO_SESSION_TTL_MS) {
      cleanupSession(session);
      macroSessions.delete(id);
    }
  }
}

/** 判断结果是否为"暂停等待弹窗处理"（session 已保存） */
function isSessionPausedResult(result: ExcelComResult): boolean {
  return !!result.details && typeof result.details === "object" && "sessionId" in result.details;
}

/**
 * 核心轮询循环：检测宏状态和弹窗。
 * interactive=true 时，检测到非 info 弹窗会返回弹窗信息并保存 session（不终止子进程）。
 * interactive=false 时，自动处理所有弹窗（保持原行为）。
 */
async function pollMacroLoop(session: MacroSession, interactive: boolean): Promise<ExcelComResult> {
  while (Date.now() < session.deadline) {
    const status = await readJsonFile<MacroRunnerStatus>(session.statusPath);
    if (status) {
      return buildMacroRunResult(status, session.macroName, session.workbookName, session.startedAt, session.collectedDialogs, session.lastDialogInspectionError, session.captureResultRange);
    }

    const inspection = await inspectExcelDialogs();
    if (inspection.error) {
      session.lastDialogInspectionError = inspection.error;
    }
    for (const rawDialog of inspection.dialogs) {
      const dialog: MacroDialogInfo = { ...rawDialog, kind: classifyMacroDialog(rawDialog) };
      const key = dialog.handle;
      const fingerprint = buildDialogFingerprint(dialog);
      const existing = session.dialogState.get(key);
      if (!existing || existing.fingerprint !== fingerprint) {
        session.dialogState.set(key, { fingerprint, firstSeen: Date.now(), handled: false, recorded: false });
        continue;
      }
      if (existing.handled || Date.now() - existing.firstSeen < MACRO_DIALOG_STABLE_MS) {
        continue;
      }
      if (!existing.recorded) {
        if (!session.recordedFingerprints.has(fingerprint)) {
          session.collectedDialogs.push({ ...dialog });
          session.recordedFingerprints.add(fingerprint);
        }
        existing.recorded = true;
      }

      const action = pickDialogAction(dialog);

      // InputBox 弹窗（action 为 null）
      if (!action) {
        if (dialog.kind === "input" && session.autoFillInputs && session.autoFillInputs.length > 0) {
          const inputText = session.autoFillInputs.shift()!;
          const filled = await fillInputBoxAndSubmit(dialog.handle, inputText);
          if (filled) {
            dialog.autoHandled = true;
            dialog.autoAction = `fill:${inputText}`;
            const dialogIndex = session.collectedDialogs.findIndex((item) => buildDialogFingerprint(item) === fingerprint);
            if (dialogIndex >= 0) {
              session.collectedDialogs[dialogIndex] = { ...session.collectedDialogs[dialogIndex], autoHandled: true, autoAction: `fill:${inputText}` };
            }
            existing.handled = true;
            await sleep(500);
          }
          continue;
        }
        // 无 autoFillInputs：interactive 模式返回弹窗信息让 AI 处理
        if (interactive) {
          macroSessions.set(session.sessionId, session);
          return {
            success: false,
            message: `宏执行期间检测到 InputBox 弹窗，宏已暂停。请调用 excel_fill_dialog 填充文本或 excel_click_dialog 点击按钮，然后调用 excel_resume_macro 恢复执行。`,
            details: {
              sessionId: session.sessionId,
              macroName: session.macroName,
              workbook: session.workbookName,
              dialogs: session.collectedDialogs,
              pendingDialog: enrichDialogInfo(dialog),
            },
          };
        }
        continue;
      }

      // 非 input 类型：interactive 模式下所有弹窗都返回让 AI 判断（包括 info 类型）
      if (interactive) {
        macroSessions.set(session.sessionId, session);
        const kindHint = dialog.kind === "vb_runtime_error"
          ? "VBA 运行时错误/编译错误"
          : dialog.kind === "error"
            ? "错误弹窗"
            : dialog.kind === "confirmation"
              ? "确认弹窗"
              : dialog.kind === "info"
                ? "信息弹窗"
                : "未知弹窗";
        return {
          success: false,
          message: `宏执行期间检测到${kindHint}：${summarizeDialog(dialog)}。请根据弹窗内容调用 excel_click_dialog 处理（如点击 结束/调试/确定/取消 等），或调用 excel_fill_dialog 填充输入框，然后调用 excel_resume_macro 恢复执行。`,
          details: {
            sessionId: session.sessionId,
            macroName: session.macroName,
            workbook: session.workbookName,
            kind: dialog.kind,
            dialogs: session.collectedDialogs,
            pendingDialog: enrichDialogInfo(dialog),
          },
        };
      }

      // 非 interactive 模式：自动处理所有弹窗（原行为）
      const handled = await invokeExcelDialogButton(dialog.handle, action.buttons);
      if (!handled) continue;

      dialog.autoHandled = true;
      dialog.autoAction = action.actionLabel;
      const dialogIndex = session.collectedDialogs.findIndex((item) => buildDialogFingerprint(item) === fingerprint);
      if (dialogIndex >= 0) {
        session.collectedDialogs[dialogIndex] = { ...session.collectedDialogs[dialogIndex], autoHandled: true, autoAction: action.actionLabel };
      } else {
        session.collectedDialogs.push(dialog);
        session.recordedFingerprints.add(fingerprint);
      }
      existing.handled = true;

      if (action.terminal) {
        await sleep(900);
      }
    }

    if (session.child.exitCode !== null) {
      break;
    }
    await sleep(MACRO_DIALOG_POLL_MS);
  }

  // 子进程结束
  if (session.child.exitCode !== null) {
    const finalStatus = await waitForRunnerStatus(session.statusPath, MACRO_RUNNER_STATUS_GRACE_MS);
    if (finalStatus) {
      return buildMacroRunResult(finalStatus, session.macroName, session.workbookName, session.startedAt, session.collectedDialogs, session.lastDialogInspectionError, session.captureResultRange);
    }
    const finishedDetails = {
      macroName: session.macroName,
      workbook: session.workbookName,
      durationMs: Date.now() - session.startedAt,
      dialogs: session.collectedDialogs,
      dialogInspectionError: session.lastDialogInspectionError,
      exitCode: session.child.exitCode,
      statusMissing: true,
    };
    if (session.child.exitCode === 0) {
      const autoHandledCount = session.collectedDialogs.filter((dialog) => dialog.autoHandled).length;
      return {
        success: true,
        message: autoHandledCount > 0
          ? `宏 ${session.macroName} 执行完成（已自动处理 ${autoHandledCount} 个弹窗）`
          : `宏 ${session.macroName} 执行完成`,
        details: finishedDetails,
      };
    }
    return { success: false, message: `宏执行进程已结束，但未返回可解析的状态（exit code: ${session.child.exitCode}）。`, details: finishedDetails };
  }

  // 超时
  const timeoutDetails = {
    macroName: session.macroName,
    workbook: session.workbookName,
    durationMs: Date.now() - session.startedAt,
    dialogs: session.collectedDialogs,
    dialogInspectionError: session.lastDialogInspectionError,
  };
  const timeoutMessage = session.lastDialogInspectionError
    ? `宏执行期间弹窗探测失败：${session.lastDialogInspectionError}`
    : session.collectedDialogs.length > 0
      ? `宏执行超时，期间检测到并处理了弹窗。最后一个弹窗：${summarizeDialog(session.collectedDialogs[session.collectedDialogs.length - 1])}`
      : "宏执行超时，Excel 可能仍在运行宏或弹出了未能自动处理的对话框。";
  return { success: false, message: timeoutMessage, details: timeoutDetails };
}

export async function runMacroWithDialogHandling(
  filePath: string,
  macroName: string,
  options: { timeoutMs?: number; captureResultRange?: string; autoFillInputs?: string[]; interactive?: boolean } = {}
): Promise<ExcelComResult> {
  const normalizedMacroName = macroName.trim();
  if (!normalizedMacroName) {
    return { success: false, message: "macroName 不能为空。" };
  }

  const wbName = basename(filePath);
  const preCheck = await ensureExcelRunning(wbName);
  if (preCheck) return preCheck;

  cleanupExpiredSessions();

  const interactive = options.interactive ?? true;
  const timeoutMs = options.timeoutMs ?? 45000;
  const sessionId = randomBytes(8).toString("hex");
  const statusPath = join(tmpdir(), `vba_macro_status_${sessionId}.json`);
  const runnerPath = join(tmpdir(), `vba_macro_runner_${sessionId}.ps1`);
  const bom = "\uFEFF";
  const utf8Header = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
`;
  await writeFile(
    runnerPath,
    bom + utf8Header + buildMacroRunnerScript({ filePath, workbookName: wbName, macroName: normalizedMacroName, statusPath }),
    "utf-8"
  );

  const child = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", runnerPath], {
    windowsHide: true,
    stdio: "ignore",
  });

  const session: MacroSession = {
    sessionId,
    child,
    statusPath,
    runnerPath,
    startedAt: Date.now(),
    deadline: Date.now() + timeoutMs,
    createdAt: Date.now(),
    macroName: normalizedMacroName,
    workbookName: wbName,
    collectedDialogs: [],
    dialogState: new Map(),
    recordedFingerprints: new Set(),
    captureResultRange: options.captureResultRange,
    autoFillInputs: options.autoFillInputs ? [...options.autoFillInputs] : undefined,
  };

  let sessionSaved = false;
  try {
    const result = await pollMacroLoop(session, interactive);
    sessionSaved = isSessionPausedResult(result);
    return result;
  } finally {
    if (!sessionSaved) {
      cleanupSession(session);
    }
  }
}

/**
 * 恢复暂停的宏执行。
 * AI 处理完弹窗后调用此函数，继续等待宏完成或检测到新弹窗。
 */
export async function resumeMacroRun(
  sessionId: string,
  options: { extendTimeoutMs?: number } = {}
): Promise<ExcelComResult> {
  const session = macroSessions.get(sessionId);
  if (!session) {
    return { success: false, message: `会话 ${sessionId} 不存在或已过期。请重新调用 excel_run_macro。` };
  }

  // 从 Map 移除（pollMacroLoop 会在需要时重新保存）
  macroSessions.delete(sessionId);

  // 可选延长超时
  if (options.extendTimeoutMs) {
    session.deadline = Date.now() + options.extendTimeoutMs;
  }

  let sessionSaved = false;
  try {
    const result = await pollMacroLoop(session, true);
    sessionSaved = isSessionPausedResult(result);
    return result;
  } finally {
    if (!sessionSaved) {
      cleanupSession(session);
    }
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
    // 这里不直接读取，因为读取需要在主 PowerShell 中执行；把需求留给调用方或后续实现
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
