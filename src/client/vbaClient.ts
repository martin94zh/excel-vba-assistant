/**
 * Excel VBA Assistant - VBA 客户端
 *
 * 通过 PowerShell 调用 Excel COM Automation 操作已打开的工作簿，
 * 实现 VBAProject 的读写、同步、宏执行与单元格读取。
 *
 * 本地同步目录结构（与 VBE 资源管理器原生分类保持一致）：
 *   syncDirectory/
 *   ├─ workbook.json                            清单
 *   ├─ Microsoft Excel 对象/ThisWorkbook.wbk    工作簿对象模块
 *   ├─ Microsoft Excel 对象/Sheet1.wks          工作表对象模块
 *   ├─ 模块/Module1.bas                         标准模块
 *   ├─ 类模块/Class1.cls                        类模块
 *   └─ 窗体/UserForm1.frm                       窗体（含 .frx）
 */
import { basename, join, resolve } from "path";
import { mkdir, writeFile, readFile, readdir } from "fs/promises";

import {
  ensureExcelRunning,
  escapePowerShellSingleQuoted,
  runPowerShell,
  sleep,
  type ExcelComResult,
} from "../runtime/powershell";
import { runMacroWithDialogHandling, listExcelDialogs, clickExcelDialog, fillDialogInput } from "./vbaMacroRunner";

/** 单元格格式 */
export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // 十六进制颜色，如 #FF0000
  backgroundColor?: string;
  fontSize?: number;
  numberFormat?: string;
  horizontalAlignment?: "Left" | "Center" | "Right";
}

/** 单元格值类型 */
export type CellValue = string | number | boolean | null;

/** 组件类型 */
export type ComponentType =
  | "document"
  | "worksheet"
  | "standardModule"
  | "classModule"
  | "userForm";

/** workbook.json 中的组件条目 */
export interface WorkbookComponentEntry {
  name: string;
  type: ComponentType;
  file: string;
}

/** workbook.json 结构 */
export interface WorkbookManifest {
  workbookName: string;
  workbookPath: string;
  lastSyncAt: string;
  components: WorkbookComponentEntry[];
}

/** VBA 组件类型常量（vbext_ComponentType） */
const CT_STD_MODULE = 1;
const CT_CLASS_MODULE = 2;
const CT_MSFORM = 3;
const CT_DOCUMENT = 100;

function componentTypeName(typeCode: number): ComponentType {
  switch (typeCode) {
    case CT_STD_MODULE: return "standardModule";
    case CT_CLASS_MODULE: return "classModule";
    case CT_MSFORM: return "userForm";
    case CT_DOCUMENT: return "document";
    default: return "document";
  }
}

function escapeNewlineForPs(value: string): string {
  return value.replace(/'/g, "''");
}

export class VbaClient {
  constructor(private readonly filePath: string) {}

  /** 检查 Excel 是否可访问 VBAProject，返回错误信息或 null */
  async checkAccess(): Promise<string | null> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck.message;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { Write-Output "WB_NOT_FOUND"; exit }
    $null = $wb.VBProject.VBComponents.Count
    Write-Output "OK"
} catch {
    $msg = [string]$_.Exception.Message
    if ($msg -match "programmatic access|不信任|access.*vba") {
        Write-Output "VBA_ACCESS_DENIED"
    } else {
        Write-Output "ERROR:$msg"
    }
}
`;
    const result = await runPowerShell(script);
    const output = (result.output || "").trim();
    if (output === "OK") return null;
    if (output === "VBA_ACCESS_DENIED") {
      return "无法访问 VBAProject。请在 Excel 中开启：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」。";
    }
    if (output === "WB_NOT_FOUND") {
      return `Excel 已运行，但未找到文件「${wbName}」。请在 Excel 中打开该文件后重试。`;
    }
    return output.startsWith("ERROR:") ? output.slice(6) : "Excel COM 调用失败，请检查 Excel 是否正常运行。";
  }

  /** 尝试在 Excel 中打开目标工作簿（单试一次） */
  private async tryOpenWorkbookInExcelOnce(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const absPath = this.filePath.replace(/\//g, "\\");
    const script = `
$ErrorActionPreference = "Stop"
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class JrExcelPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
function Jr-OpenWorkbook {
  try {
    $excel = $null
    $started = $false
    try {
        $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $started = $true
    }
    $wb = $null
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break }
    }
    if ($wb -eq $null) {
        $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(absPath)}')
    }
    $excel.Visible = $true
    $pidValue = [uint32]0
    $hwnd = $excel.Hwnd
    $hwndPtr = [IntPtr]::new([long]$hwnd)
    [void][JrExcelPid]::GetWindowThreadProcessId($hwndPtr, [ref]$pidValue)
    $payload = @{ success = $true; workbookName = '${escapePowerShellSingleQuoted(wbName)}'; startedExcel = $started; message = "已在 Excel 中打开文件"; pid = [int]$pidValue }
    Write-Output (ConvertTo-Json $payload -Compress)
  } catch {
    $payload = @{ success = $false; error = [string]$_.Exception.Message; stack = [string]$_.ScriptStackTrace }
    Write-Output (ConvertTo-Json $payload -Compress)
  }
}
Jr-OpenWorkbook
`;
    const result = await runPowerShell(script);
    if (result.success && result.output) {
      try {
        const parsed = JSON.parse(result.output);
        if (parsed.success) {
          return { success: true, message: parsed.message || "已打开", output: result.output, details: { pid: parsed.pid } };
        }
        const reason = parsed.error || "打开 Excel 失败";
        return {
          success: false,
          message: `${reason}。请先打开 Excel 并加载文件「${wbName}」后重试。`,
          output: result.output,
        };
      } catch {
        return result;
      }
    }
    return result;
  }

  /** 尝试在 Excel 中打开目标工作簿，失败时自动重试 3 次。Excel 未运行时会自动启动并显示窗口。返回的 details.pid 为 Excel 进程 ID。 */
  async openWorkbookInExcel(): Promise<ExcelComResult> {
    const maxRetries = 3;
    let lastResult: ExcelComResult | undefined;
    for (let i = 0; i < maxRetries; i++) {
      lastResult = await this.tryOpenWorkbookInExcelOnce();
      if (lastResult.success) return lastResult;
      if (i < maxRetries - 1) {
        await sleep(1000);
      }
    }
    return lastResult!;
  }

  // ============================================================
  // 资源读取
  // ============================================================

  /** 获取工作簿资源（VBComponents 列表） */
  async getResources(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $result = @()
    foreach ($comp in $wb.VBProject.VBComponents) {
        $ct = [int]$comp.Type
        $kind = switch ($ct) { 1 { "StdModule" } 2 { "ClassModule" } 3 { "UserForm" } 100 { if ($comp.Name -eq "ThisWorkbook") { "ThisWorkbook" } else { "Worksheet" } } default { "Unknown" } }
        $hasCode = $false
        try { if ($comp.CodeModule.CountOfLines -gt 0) { $hasCode = $true } } catch {}
        $result += [pscustomobject]@{
            name = [string]$comp.Name
            type = $kind
            codeLines = [int]$comp.CodeModule.CountOfLines
            hasCode = $hasCode
        }
    }
    $payload = @{ workbookName = '${escapePowerShellSingleQuoted(wbName)}'; workbookPath = '${escapePowerShellSingleQuoted(this.filePath)}'; items = $result }
    Write-Output (ConvertTo-Json $payload -Depth 5 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

/** 一次性返回工作簿全景信息（工作表、VBA 资源、宏、代码行数统计） */
  async inspectWorkbook(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }

    # 工作表
    $sheets = @()
    $idx = 0
    foreach ($ws in $wb.Worksheets) {
        $idx++
        $usedRange = ""
        try { $usedRange = [string]$ws.UsedRange.Address($false, $false) } catch {}
        $sheets += [pscustomobject]@{ name = [string]$ws.Name; index = $idx; usedRange = $usedRange }
    }

    # VBA 资源
    $resources = @()
    $totalCodeLines = 0
    foreach ($comp in $wb.VBProject.VBComponents) {
        $ct = [int]$comp.Type
        $kind = switch ($ct) { 1 { "StdModule" } 2 { "ClassModule" } 3 { "UserForm" } 100 { if ($comp.Name -eq "ThisWorkbook") { "ThisWorkbook" } else { "Worksheet" } } default { "Unknown" } }
        $lines = 0
        try { $lines = [int]$comp.CodeModule.CountOfLines } catch {}
        $totalCodeLines += $lines
        $resources += [pscustomobject]@{ name = [string]$comp.Name; type = $kind; codeLines = $lines }
    }

    # 宏列表
    $macros = @()
    foreach ($comp in $wb.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        $lines = $cm.CountOfLines
        for ($i = 1; $i -le $lines; $i++) {
            $decl = $cm.Lines($i, 1)
            if ($decl -match '^\s*(Public\s+|Private\s+)?(Sub|Function)\s+(\w+)') {
                $procName = $matches[3]
                $scope = if ($matches[1]) { $matches[1].Trim() } else { "Public" }
                $kind = $matches[2]
                $fullName = if ($comp.Type -eq 100) { "$($comp.Name).$procName" } else { "$($comp.Name).$procName" }
                $macros += [pscustomobject]@{ component = [string]$comp.Name; name = $procName; fullName = $fullName; kind = $kind; scope = $scope }
            }
        }
    }

    $payload = @{
        workbookName = '${escapePowerShellSingleQuoted(wbName)}'
        workbookPath = '${escapePowerShellSingleQuoted(this.filePath)}'
        sheets = $sheets
        resources = $resources
        totalCodeLines = $totalCodeLines
        macros = $macros
    }
    Write-Output (ConvertTo-Json $payload -Depth 10 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 读取所有 VBA 组件代码 */
  async getAllComponentCode(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $items = @()
    foreach ($comp in $wb.VBProject.VBComponents) {
        $ct = [int]$comp.Type
        $kind = switch ($ct) { 1 { "StdModule" } 2 { "ClassModule" } 3 { "UserForm" } 100 { if ($comp.Name -eq "ThisWorkbook") { "ThisWorkbook" } else { "Worksheet" } } default { "Unknown" } }
        $code = ""
        try {
            if ($comp.CodeModule.CountOfLines -gt 0) {
                $code = [string]$comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
            }
        } catch {}
        $items += [pscustomobject]@{ name = [string]$comp.Name; type = $kind; codeLines = [int]$comp.CodeModule.CountOfLines; code = $code }
    }
    $payload = @{ workbookName = '${escapePowerShellSingleQuoted(wbName)}'; items = $items }
    Write-Output (ConvertTo-Json $payload -Depth 10 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 读取单个组件代码 */
  async getComponentCode(componentName: string): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $comp = $null
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq '${escapePowerShellSingleQuoted(componentName)}') { $comp = $c; break } }
    if ($comp -eq $null) { Write-Output '{"error":"未找到组件 ${escapePowerShellSingleQuoted(componentName)}"}'; exit }
    $ct = [int]$comp.Type
    $typeName = switch ($ct) { 1 { "standardModule" } 2 { "classModule" } 3 { "userForm" } 100 { "document" } default { "unknown" } }
    $code = ""
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -gt 0) { $code = [string]$cm.Lines(1, $cm.CountOfLines) }
    $codeBytes = [System.Text.Encoding]::UTF8.GetBytes($code)
    $codeBase64 = [System.Convert]::ToBase64String($codeBytes)
    $payload = @{ workbookName = '${escapePowerShellSingleQuoted(wbName)}'; componentName = '${escapePowerShellSingleQuoted(componentName)}'; componentType = $typeName; codeBase64 = $codeBase64 }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    const result = await runPowerShell(script);
    if (result.success && result.output) {
      try {
        const parsed = JSON.parse(result.output);
        if (parsed.codeBase64) {
          parsed.code = Buffer.from(parsed.codeBase64, "base64").toString("utf-8");
          delete parsed.codeBase64;
        }
        return { ...result, output: JSON.stringify(parsed) };
      } catch {
        return result;
      }
    }
    return result;
  }

  /** 获取当前 VBE 所有组件代码的 SHA256 checksum，用于检测 VBE 是否有变更 */
  async getVbeCodeChecksum(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $sb = New-Object System.Text.StringBuilder
    $comps = $wb.VBProject.VBComponents | Sort-Object Name
    foreach ($comp in $comps) {
        $cm = $comp.CodeModule
        $code = ""
        if ($cm.CountOfLines -gt 0) { $code = [string]$cm.Lines(1, $cm.CountOfLines) }
        [void]$sb.AppendLine([string]$comp.Name)
        [void]$sb.AppendLine($code)
    }
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $hash = $sha256.ComputeHash($bytes)
    $hashHex = [BitConverter]::ToString($hash) -replace "-", ""
    Write-Output $hashHex
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  // ============================================================
  // 同步：VBE → 本地
  // ============================================================

  /** VBE → 本地：导出所有 VBA 组件到本地目录并写入 workbook.json */
  async syncVbeToLocal(localDir: string): Promise<ExcelComResult> {
    const absDir = resolve(localDir);
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;

    // 确保目录结构（与 VBE 资源管理器原生分类保持一致）
    await mkdir(join(absDir, "Microsoft Excel 对象"), { recursive: true });
    await mkdir(join(absDir, "模块"), { recursive: true });
    await mkdir(join(absDir, "类模块"), { recursive: true });
    await mkdir(join(absDir, "窗体"), { recursive: true });

    const absPath = this.filePath.replace(/\//g, "\\");
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(absPath)}') }

    $baseDir = '${escapePowerShellSingleQuoted(absDir.replace(/\//g, "\\"))}'
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $gbk = [System.Text.Encoding]::GetEncoding("GBK")
    $components = @()

    function Write-ComponentFile([string]$folderPath, [string]$fileName, [string]$ext, $comp, [int]$compType, [string]$relFile) {
        $targetFile = Join-Path $folderPath ($fileName + $ext)
        $written = $false
        if ($compType -eq 3) {
            $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("vbe_export_" + [System.Guid]::NewGuid().ToString("N").Substring(0,8))
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            $tmpFrm = Join-Path $tmpDir ($fileName + ".frm")
            $comp.Export($tmpFrm)
            $rawBytes = [System.IO.File]::ReadAllBytes($tmpFrm)
            $content = $gbk.GetString($rawBytes)

            # 仅在 .frm 内容变化时写入
            $frmChanged = $true
            if (Test-Path $targetFile) {
                $existingBytes = [System.IO.File]::ReadAllBytes($targetFile)
                $existingContent = $utf8NoBom.GetString($existingBytes)
                if ($existingContent -eq $content) { $frmChanged = $false }
            }
            if ($frmChanged) {
                [System.IO.File]::WriteAllText($targetFile, $content, $utf8NoBom)
                $written = $true
            }

            # 仅在 .frx 内容变化时复制
            $tmpFrx = Join-Path $tmpDir ($fileName + ".frx")
            $targetFrx = Join-Path $folderPath ($fileName + ".frx")
            if (Test-Path $tmpFrx) {
                $frxChanged = $true
                if (Test-Path $targetFrx) {
                    $existingFrxBytes = [System.IO.File]::ReadAllBytes($targetFrx)
                    $newFrxBytes = [System.IO.File]::ReadAllBytes($tmpFrx)
                    if ($existingFrxBytes.Length -eq $newFrxBytes.Length) {
                        $same = $true
                        for ($i = 0; $i -lt $existingFrxBytes.Length; $i++) {
                            if ($existingFrxBytes[$i] -ne $newFrxBytes[$i]) { $same = $false; break }
                        }
                        if ($same) { $frxChanged = $false }
                    }
                }
                if ($frxChanged) {
                    Copy-Item $tmpFrx $targetFrx -Force
                    $written = $true
                }
            }
            Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            $tmpExport = [System.IO.Path]::GetTempFileName() + $ext
            $comp.Export($tmpExport)
            $rawBytes = [System.IO.File]::ReadAllBytes($tmpExport)
            $content = $gbk.GetString($rawBytes)
            Remove-Item $tmpExport -ErrorAction SilentlyContinue

            # 仅在内容变化时写入，避免触发本地 → VBE 自动同步循环
            $contentChanged = $true
            if (Test-Path $targetFile) {
                $existingBytes = [System.IO.File]::ReadAllBytes($targetFile)
                $existingContent = $utf8NoBom.GetString($existingBytes)
                if ($existingContent -eq $content) { $contentChanged = $false }
            }
            if ($contentChanged) {
                [System.IO.File]::WriteAllText($targetFile, $content, $utf8NoBom)
                $written = $true
            }
        }
        if ($written) { Write-Output "同步: $relFile" }
        else { Write-Output "未变: $relFile" }
        return $targetFile
    }

    foreach ($comp in $wb.VBProject.VBComponents) {
        $ct = [int]$comp.Type
        $compName = [string]$comp.Name
        switch ($ct) {
            1 {
                $folder = Join-Path $baseDir "模块"
                $ext = ".bas"
                $relFile = "模块/$compName.bas"
                $typeName = "standardModule"
            }
            2 {
                $folder = Join-Path $baseDir "类模块"
                $ext = ".cls"
                $relFile = "类模块/$compName.cls"
                $typeName = "classModule"
            }
            3 {
                $folder = Join-Path $baseDir "窗体"
                $ext = ".frm"
                $relFile = "窗体/$compName.frm"
                $typeName = "userForm"
            }
            100 {
                $folder = Join-Path $baseDir "Microsoft Excel 对象"
                if ($compName -eq "ThisWorkbook") {
                    $ext = ".wbk"
                    $relFile = "Microsoft Excel 对象/ThisWorkbook.wbk"
                    $typeName = "document"
                } else {
                    $ext = ".wks"
                    $relFile = "Microsoft Excel 对象/$compName.wks"
                    $typeName = "worksheet"
                }
            }
            default { continue }
        }
        if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $null = Write-ComponentFile $folder $compName $ext $comp $ct $relFile
        $components += [pscustomobject]@{ name = $compName; type = $typeName; file = $relFile }
    }

    # 清理本地存在但 VBE 中已不存在的文件（删除同步：VBE → 本地）
    $expectedFiles = @{}
    foreach ($c in $components) {
        $expectedFiles[$c.file.ToLower()] = $true
        if ($c.type -eq "userForm") {
            $frxFile = $c.file -replace '\.frm$', '.frx'
            $expectedFiles[$frxFile.ToLower()] = $true
        }
    }

    function Remove-OrphanLocalFiles([string]$dir, [string]$filter) {
        if (-not (Test-Path $dir)) { return 0 }
        $removed = 0
        foreach ($f in (Get-ChildItem -Path $dir -Filter $filter -File)) {
            $rel = $f.FullName.Substring($baseDir.Length + 1).Replace("\\", "/")
            if (-not $expectedFiles.ContainsKey($rel.ToLower())) {
                Remove-Item $f.FullName -Force
                Write-Output "删除: $rel"
                $removed++
            }
        }
        return $removed
    }

    $deleted = 0
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "模块") "*.bas"
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "类模块") "*.cls"
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "窗体") "*.frm"
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "窗体") "*.frx"
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "Microsoft Excel 对象") "*.wks"
    $deleted += Remove-OrphanLocalFiles (Join-Path $baseDir "Microsoft Excel 对象") "*.wbk"

    # 清理旧版 Excel 对象目录和文件（已迁移到 .wks / .wbk），避免资源管理器中显示重复
    $oldWorksheetDir = Join-Path $baseDir "工作表"
    if (Test-Path $oldWorksheetDir) {
        Remove-Item $oldWorksheetDir -Recurse -Force
        Write-Output "清理旧目录: 工作表"
    }

    $excelObjDir = Join-Path $baseDir "Microsoft Excel 对象"
    if (Test-Path $excelObjDir) {
        foreach ($oldCls in (Get-ChildItem -Path $excelObjDir -Filter "*.cls" -File)) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($oldCls.Name)
            $wksPath = Join-Path $excelObjDir ($baseName + ".wks")
            $wbkPath = Join-Path $excelObjDir ($baseName + ".wbk")
            if ((Test-Path $wksPath) -or (Test-Path $wbkPath)) {
                Remove-Item $oldCls.FullName -Force
                Write-Output "清理旧文件: Microsoft Excel 对象/$($oldCls.Name)"
            }
        }
    }

    foreach ($oldRootCls in (Get-ChildItem -Path $baseDir -Filter "Sheet*.cls" -File)) {
        Remove-Item $oldRootCls.FullName -Force
        Write-Output "清理旧文件: $($oldRootCls.Name)"
    }
    $oldThisWorkbookCls = Join-Path $baseDir "ThisWorkbook.cls"
    if (Test-Path $oldThisWorkbookCls) {
        Remove-Item $oldThisWorkbookCls -Force
        Write-Output "清理旧文件: ThisWorkbook.cls"
    }

    $manifest = @{
        workbookName = '${escapePowerShellSingleQuoted(wbName)}'
        workbookPath = '${escapePowerShellSingleQuoted(this.filePath)}'
        lastSyncAt = (Get-Date).ToString("o")
        components = $components
    }
    $manifestJson = ConvertTo-Json $manifest -Depth 5
    $manifestFile = Join-Path $baseDir "workbook.json"
    [System.IO.File]::WriteAllText($manifestFile, $manifestJson, $utf8NoBom)
    Write-Output ""
    Write-Output "同步完成: 共 $($components.Count) 个组件, 删除 $deleted 个本地文件"
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    const result = await runPowerShell(script);
    return result;
  }

  // ============================================================
  // 同步：本地 → VBE
  // ============================================================

  /** 本地 → VBE：读取本地文件覆盖写回 VBE（覆盖性写入） */
  async syncLocalToVbe(localDir: string): Promise<ExcelComResult> {
    const absDir = resolve(localDir);
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;

    const absPath = this.filePath.replace(/\//g, "\\");
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(absPath)}') }
    $vbProject = $wb.VBProject
    $baseDir = '${escapePowerShellSingleQuoted(absDir.replace(/\//g, "\\"))}'
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    function Get-CleanCode([string]$filePath) {
        $content = [System.IO.File]::ReadAllText($filePath, $utf8NoBom)
        $lines = $content -split "\`r\`n|\`r|\`n"
        $clean = @()
        $skipVersion = $false
        foreach ($l in $lines) {
            $t = $l.TrimEnd()
            if ($t -match "^VERSION ") { $skipVersion = $true; continue }
            if ($skipVersion -and $t -eq "END") { $skipVersion = $false; continue }
            if ($skipVersion) { continue }
            if ($t -match "^Attribute ") { continue }
            if ($t -match "^Begin |^End$") { continue }
            $clean += $t
        }
        return ($clean -join "\`r\`n").TrimEnd()
    }

    function Get-ComponentNameFromFile([string]$fileName) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        return $name
    }

    function Find-Component($vbProject, [string]$name, [int]$typeCode) {
        foreach ($c in $vbProject.VBComponents) {
            if ($c.Name -eq $name -and [int]$c.Type -eq $typeCode) { return $c }
        }
        return $null
    }

    function Update-ComponentCode($comp, [string]$code) {
        $cm = $comp.CodeModule
        $vbeCode = ""
        if ($cm.CountOfLines -gt 0) { $vbeCode = [string]$cm.Lines(1, $cm.CountOfLines) }
        $vbeClean = ($vbeCode -split "\`r\`n|\`r|\`n" | Where-Object { $_ -notmatch "^Attribute " }) -join "\`r\`n"
        if ($vbeClean.TrimEnd() -eq $code.TrimEnd()) { return $false }
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        if ($code.Length -gt 0) { $cm.AddFromString($code) }
        return $true
    }

    $added = 0; $updated = 0; $unchanged = 0; $skipped = 0

    # 标准模块 模块/*.bas
    $modulesDir = Join-Path $baseDir "模块"
    if (Test-Path $modulesDir) {
        foreach ($file in (Get-ChildItem -Path $modulesDir -Filter "*.bas" -File)) {
            $modName = Get-ComponentNameFromFile $file.Name
            $localCode = Get-CleanCode $file.FullName
            $exist = Find-Component $vbProject $modName 1
            if ($exist -ne $null) {
                if (Update-ComponentCode $exist $localCode) {
                    $updated++
                    Write-Output "更新: 模块/$modName.bas"
                } else {
                    $unchanged++
                    Write-Output "未变: 模块/$modName.bas"
                }
            } else {
                $newComp = $vbProject.VBComponents.Add(1)
                $newComp.Name = $modName
                if ($localCode.Length -gt 0) { $newComp.CodeModule.AddFromString($localCode) }
                $added++
                Write-Output "新增: 模块/$modName.bas"
            }
        }
    }

    # 类模块 类模块/*.cls
    $classesDir = Join-Path $baseDir "类模块"
    if (Test-Path $classesDir) {
        foreach ($file in (Get-ChildItem -Path $classesDir -Filter "*.cls" -File)) {
            $modName = Get-ComponentNameFromFile $file.Name
            $localCode = Get-CleanCode $file.FullName
            $exist = Find-Component $vbProject $modName 2
            if ($exist -ne $null) {
                if (Update-ComponentCode $exist $localCode) {
                    $updated++
                    Write-Output "更新: 类模块/$modName.cls"
                } else {
                    $unchanged++
                    Write-Output "未变: 类模块/$modName.cls"
                }
            } else {
                $newComp = $vbProject.VBComponents.Add(2)
                $newComp.Name = $modName
                if ($localCode.Length -gt 0) { $newComp.CodeModule.AddFromString($localCode) }
                $added++
                Write-Output "新增: 类模块/$modName.cls"
            }
        }
    }

    # Excel 对象模块 Microsoft Excel 对象/*.wks（工作表）/*.wbk（工作簿）（只覆盖，不新建）
    $excelObjectsDir = Join-Path $baseDir "Microsoft Excel 对象"
    if (Test-Path $excelObjectsDir) {
        foreach ($file in (Get-ChildItem -Path $excelObjectsDir -Filter "*.wks" -File)) {
            $modName = Get-ComponentNameFromFile $file.Name
            $localCode = Get-CleanCode $file.FullName
            $exist = Find-Component $vbProject $modName 100
            if ($exist -ne $null) {
                if (Update-ComponentCode $exist $localCode) {
                    $updated++
                    Write-Output "更新: Microsoft Excel 对象/$modName.wks (事件代码)"
                } else {
                    $unchanged++
                    Write-Output "未变: Microsoft Excel 对象/$modName.wks"
                }
            } else {
                $skipped++
                Write-Output "跳过: Microsoft Excel 对象/$modName.wks (VBE 中不存在，Excel 对象模块不会从本地新建)"
            }
        }
        foreach ($file in (Get-ChildItem -Path $excelObjectsDir -Filter "*.wbk" -File)) {
            $modName = Get-ComponentNameFromFile $file.Name
            $localCode = Get-CleanCode $file.FullName
            $exist = Find-Component $vbProject $modName 100
            if ($exist -ne $null) {
                if (Update-ComponentCode $exist $localCode) {
                    $updated++
                    Write-Output "更新: Microsoft Excel 对象/$modName.wbk (事件代码)"
                } else {
                    $unchanged++
                    Write-Output "未变: Microsoft Excel 对象/$modName.wbk"
                }
            } else {
                $skipped++
                Write-Output "跳过: Microsoft Excel 对象/$modName.wbk (VBE 中不存在，Excel 对象模块不会从本地新建)"
            }
        }
    }

    # 窗体 窗体/*.frm（只覆盖代码部分，不新建窗体）
    $formsDir = Join-Path $baseDir "窗体"
    if (Test-Path $formsDir) {
        foreach ($file in (Get-ChildItem -Path $formsDir -Filter "*.frm" -File)) {
            $modName = Get-ComponentNameFromFile $file.Name
            $localCode = Get-CleanCode $file.FullName
            $exist = Find-Component $vbProject $modName 3
            if ($exist -ne $null) {
                if (Update-ComponentCode $exist $localCode) {
                    $updated++
                    Write-Output "更新: 窗体/$modName.frm (事件代码)"
                } else {
                    $unchanged++
                    Write-Output "未变: 窗体/$modName.frm"
                }
            } else {
                $skipped++
                Write-Output "跳过: 窗体/$modName.frm (VBE 中不存在，请先在 VBE 中创建此窗体)"
            }
        }
    }

    # 删除同步：本地 → VBE
    # VBE 中存在但本地已不存在的标准模块/类模块/窗体会被删除
    # Excel 对象模块（Type=100）永远不参与删除，避免误删工作表/工作簿
    $localModules = @{}
    if (Test-Path $modulesDir) {
        foreach ($file in (Get-ChildItem -Path $modulesDir -Filter "*.bas" -File)) {
            $localModules[(Get-ComponentNameFromFile $file.Name).ToLower()] = $true
        }
    }
    $localClasses = @{}
    if (Test-Path $classesDir) {
        foreach ($file in (Get-ChildItem -Path $classesDir -Filter "*.cls" -File)) {
            $localClasses[(Get-ComponentNameFromFile $file.Name).ToLower()] = $true
        }
    }
    $localForms = @{}
    if (Test-Path $formsDir) {
        foreach ($file in (Get-ChildItem -Path $formsDir -Filter "*.frm" -File)) {
            $localForms[(Get-ComponentNameFromFile $file.Name).ToLower()] = $true
        }
    }

    $toRemove = @()
    $vbeDeleted = 0
    foreach ($comp in $vbProject.VBComponents) {
        $ct = [int]$comp.Type
        $name = [string]$comp.Name
        $nameLower = $name.ToLower()
        if ($ct -eq 1 -and -not $localModules.ContainsKey($nameLower)) {
            $toRemove += $comp
        } elseif ($ct -eq 2 -and -not $localClasses.ContainsKey($nameLower)) {
            $toRemove += $comp
        } elseif ($ct -eq 3 -and -not $localForms.ContainsKey($nameLower)) {
            $toRemove += $comp
        }
    }
    foreach ($comp in $toRemove) {
        $name = [string]$comp.Name
        $ct = [int]$comp.Type
        $vbProject.VBComponents.Remove($comp)
        $vbeDeleted++
        if ($ct -eq 1) { Write-Output "删除: 模块/$name.bas" }
        elseif ($ct -eq 2) { Write-Output "删除: 类模块/$name.cls" }
        elseif ($ct -eq 3) { Write-Output "删除: 窗体/$name.frm" }
    }

    $wb.Save()

    # 更新 workbook.json 以反映当前 VBE 状态
    $currentComponents = @()
    foreach ($comp in $vbProject.VBComponents) {
        $ct = [int]$comp.Type
        $compName = [string]$comp.Name
        $typeName = switch ($ct) { 1 { "standardModule" } 2 { "classModule" } 3 { "userForm" } 100 { if ($compName -eq "ThisWorkbook") { "document" } else { "worksheet" } } default { "unknown" } }
        $relFile = switch ($ct) {
            1 { "模块/$compName.bas" }
            2 { "类模块/$compName.cls" }
            3 { "窗体/$compName.frm" }
            100 { if ($compName -eq "ThisWorkbook") { "Microsoft Excel 对象/ThisWorkbook.wbk" } else { "Microsoft Excel 对象/$compName.wks" } }
            default { "" }
        }
        if ($relFile -ne "") {
            $currentComponents += [pscustomobject]@{ name = $compName; type = $typeName; file = $relFile }
        }
    }
    $manifest = @{
        workbookName = [string]$wb.Name
        workbookPath = '${escapePowerShellSingleQuoted(this.filePath)}'
        lastSyncAt = (Get-Date).ToString("o")
        components = $currentComponents
    }
    $manifestFile = Join-Path $baseDir "workbook.json"
    [System.IO.File]::WriteAllText($manifestFile, (ConvertTo-Json $manifest -Depth 5), $utf8NoBom)

    Write-Output ""
    Write-Output "同步完成: 新增 $added / 更新 $updated / 跳过 $skipped / 未变 $unchanged / 删除 $vbeDeleted"
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    const result = await runPowerShell(script);
    if (result.success) {
      // 更新 workbook.json 的 lastSyncAt
      await this.updateManifestLastSync(absDir, "local-to-vbe");
    }
    return result;
  }

  /** 更新 workbook.json 的同步时间 */
  private async updateManifestLastSync(localDir: string, _direction: string): Promise<void> {
    try {
      const manifestPath = join(localDir, "workbook.json");
      const raw = await readFile(manifestPath, "utf-8");
      const manifest = JSON.parse(raw) as WorkbookManifest;
      manifest.lastSyncAt = new Date().toISOString();
      await writeFile(manifestPath, JSON.stringify(manifest, null, 2), "utf-8");
    } catch {
      // 清单不存在时忽略
    }
  }

  // ============================================================
  // 宏执行
  // ============================================================

  /** 列出所有宏（解析各组件中的 Sub/Function） */
  async listMacros(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $macros = @()
    foreach ($comp in $wb.VBProject.VBComponents) {
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -eq 0) { continue }
        $code = [string]$cm.Lines(1, $cm.CountOfLines)
        $lines = $code -split "\`r\`n|\`r|\`n"
        foreach ($l in $lines) {
            if ($l -match "^\\s*(?:Public\\s+|Private\\s+)?(?:Static\\s+)?(Sub|Function)\\s+([A-Za-z_][A-Za-z0-9_]*)") {
                $procName = $Matches[2]
                $macros += [pscustomobject]@{ name = "$($comp.Name).$procName"; module = [string]$comp.Name; procedure = $procName }
            }
        }
    }
    $payload = @{ macros = $macros }
    Write-Output (ConvertTo-Json $payload -Depth 5 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 运行无参数宏 */
  async runMacro(
    macroName: string,
    options: { timeoutMs?: number; captureResultRange?: string; silent?: boolean } = {}
  ): Promise<ExcelComResult> {
    return runMacroWithDialogHandling(this.filePath, macroName, {
      timeoutMs: options.timeoutMs,
      captureResultRange: options.captureResultRange,
    });
  }

  /** 列出当前 Excel/VBA 弹窗 */
  async listDialogs(): Promise<ExcelComResult> {
    return listExcelDialogs();
  }

  /** 点击指定弹窗的按钮 */
  async clickDialog(handle: string, action?: string, buttonText?: string): Promise<ExcelComResult> {
    return clickExcelDialog(handle, action, buttonText);
  }

  /** 向指定弹窗的输入框写入文本 */
  async fillDialog(handle: string, text: string, submit = false): Promise<ExcelComResult> {
    return fillDialogInput(handle, text, submit);
  }

  // ============================================================
  // 工作表 / 单元格读取
  // ============================================================

  /** 列出工作表 */
  async listSheets(): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $sheets = @()
    $idx = 0
    foreach ($ws in $wb.Worksheets) {
        $idx++
        $usedRange = ""
        try { $usedRange = [string]$ws.UsedRange.Address($false, $false) } catch {}
        $sheets += [pscustomobject]@{ name = [string]$ws.Name; index = $idx; usedRange = $usedRange }
    }
    $payload = @{ sheets = $sheets }
    Write-Output (ConvertTo-Json $payload -Depth 5 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 读取 UsedRange；sheetName 为空时使用第一个工作表 */
  async readUsedRange(sheetName: string | undefined): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $rng = $ws.UsedRange
    $address = [string]$rng.Address($false, $false)
    $rows = [int]$rng.Rows.Count
    $cols = [int]$rng.Columns.Count
    $values = @()
    for ($r = 1; $r -le $rows; $r++) {
        $row = @()
        for ($c = 1; $c -le $cols; $c++) {
            $val = $rng.Cells.Item($r, $c).Value2
            if ($val -eq $null) { $val = "" }
            $row += $val
        }
        $values += ,($row)
    }
    $actualSheet = [string]$ws.Name
    $payload = @{ sheetName = $actualSheet; address = $address; values = $values }
    Write-Output (ConvertTo-Json $payload -Depth 10 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  /** 读取指定 Range；sheetName 为空时使用第一个工作表 */
  async readRange(sheetName: string | undefined, address: string): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $rng = $ws.Range('${escapePowerShellSingleQuoted(address)}')
    $rows = [int]$rng.Rows.Count
    $cols = [int]$rng.Columns.Count
    $values = @()
    for ($r = 1; $r -le $rows; $r++) {
        $row = @()
        for ($c = 1; $c -le $cols; $c++) {
            $val = $rng.Cells.Item($r, $c).Value2
            if ($val -eq $null) { $val = "" }
            $row += $val
        }
        $values += ,($row)
    }
    $actualSheet = [string]$ws.Name
    $payload = @{ sheetName = $actualSheet; address = '${escapePowerShellSingleQuoted(address)}'; values = $values }
    Write-Output (ConvertTo-Json $payload -Depth 10 -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  // ============================================================
  // Excel 工作簿 / 工作表 / 单元格 编辑（需求 1）
  // ============================================================

  /** 新建工作表 */
  async createSheet(name: string, options?: { before?: string; after?: string }): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const beforeLit = options?.before ? `'${escapePowerShellSingleQuoted(options.before)}'` : "$null";
    const afterLit = options?.after ? `'${escapePowerShellSingleQuoted(options.after)}'` : "$null";
    const nameLit = escapePowerShellSingleQuoted(name);
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $before = if ($beforeLit -ne "$null") { $wb.Sheets.Item($beforeLit) } else { $null }
    $after  = if ($afterLit -ne "$null") { $wb.Sheets.Item($afterLit) } else { $null }
    $ws = $wb.Sheets.Add($before, $after)
    $ws.Name = '${nameLit}'
    $wb.Save()
    $payload = @{ success = $true; sheetName = '${nameLit}'; message = "工作表 ${nameLit} 已创建" }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$beforeLit/g, beforeLit).replace(/\$afterLit/g, afterLit);
    return runPowerShell(script);
  }

  /** 删除工作表 */
  async deleteSheet(name: string): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = $wb.Sheets.Item('${escapePowerShellSingleQuoted(name)}')
    $ws.Delete()
    $wb.Save()
    $payload = @{ success = $true; sheetName = '${escapePowerShellSingleQuoted(name)}'; message = "工作表已删除" }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 重命名工作表 */
  async renameSheet(oldName: string, newName: string): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = $wb.Sheets.Item('${escapePowerShellSingleQuoted(oldName)}')
    $ws.Name = '${escapePowerShellSingleQuoted(newName)}'
    $wb.Save()
    $payload = @{ success = $true; oldName = '${escapePowerShellSingleQuoted(oldName)}'; newName = '${escapePowerShellSingleQuoted(newName)}' }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }

  /** 设置单个单元格值；sheetName 为空时使用第一个工作表 */
  async setCellValue(sheetName: string | undefined, address: string, value: CellValue): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const valueLit = value === null ? "$null" : `'${escapePowerShellSingleQuoted(String(value))}'`;
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $ws.Range('${escapePowerShellSingleQuoted(address)}').Value2 = $valueLit
    $wb.Save()
    $actualSheet = [string]$ws.Name
    $payload = @{ success = $true; sheetName = $actualSheet; address = '${escapePowerShellSingleQuoted(address)}' }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$valueLit/g, valueLit).replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  /** 批量设置 Range 值，values 为二维数组；sheetName 为空时使用第一个工作表 */
  async setRangeValues(sheetName: string | undefined, startAddress: string, values: CellValue[][]): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const rows = values.length;
    const cols = rows > 0 ? values[0].length : 0;
    if (rows === 0 || cols === 0) {
      return { success: false, message: "values 不能为空数组" };
    }
    // 将二维数组序列化为 JSON，PowerShell 中解析
    const jsonValues = JSON.stringify(values);
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $start = $ws.Range('${escapePowerShellSingleQuoted(startAddress)}')
    $endCell = $ws.Cells.Item($start.Row + ${rows - 1}, $start.Column + ${cols - 1})
    $rng = $ws.Range($start, $endCell)
    $values = ConvertFrom-Json '${escapePowerShellSingleQuoted(jsonValues)}'
    $arr = New-Object 'object[,]' ${rows}, ${cols}
    for ($r = 0; $r -lt ${rows}; $r++) {
        for ($c = 0; $c -lt ${cols}; $c++) {
            $v = $values[$r][$c]
            if ($v -eq $null) { $v = "" }
            $arr[$r, $c] = $v
        }
    }
    $rng.Value2 = $arr
    $wb.Save()
    $actualSheet = [string]$ws.Name
    $payload = @{ success = $true; sheetName = $actualSheet; address = [string]$rng.Address($false, $false) }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  /** 清空指定 Range；sheetName 为空时使用第一个工作表 */
  async clearRange(sheetName: string | undefined, address: string): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $ws.Range('${escapePowerShellSingleQuoted(address)}').ClearContents()
    $wb.Save()
    $actualSheet = [string]$ws.Name
    $payload = @{ success = $true; sheetName = $actualSheet; address = '${escapePowerShellSingleQuoted(address)}' }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  /** 设置单元格格式；sheetName 为空时使用第一个工作表 */
  async setCellFormat(sheetName: string | undefined, address: string, format: CellFormat): Promise<ExcelComResult> {
    const wbName = basename(this.filePath);
    const preCheck = await ensureExcelRunning(wbName);
    if (preCheck) return preCheck;
    const fmt = format;
    const sheetLit = sheetName ? `'${escapePowerShellSingleQuoted(sheetName)}'` : "$null";
    const psLines: string[] = [];
    if (fmt.bold !== undefined) psLines.push(`$rng.Font.Bold = ${fmt.bold ? "$true" : "$false"}`);
    if (fmt.italic !== undefined) psLines.push(`$rng.Font.Italic = ${fmt.italic ? "$true" : "$false"}`);
    if (fmt.underline !== undefined) psLines.push(`$rng.Font.Underline = ${fmt.underline ? "$true" : "$false"}`);
    if (fmt.color) psLines.push(`$rng.Font.Color = [System.Drawing.ColorTranslator]::FromHtml('${escapePowerShellSingleQuoted(fmt.color)}')`);
    if (fmt.backgroundColor) psLines.push(`$rng.Interior.Color = [System.Drawing.ColorTranslator]::FromHtml('${escapePowerShellSingleQuoted(fmt.backgroundColor)}')`);
    if (fmt.fontSize) psLines.push(`$rng.Font.Size = ${fmt.fontSize}`);
    if (fmt.numberFormat) psLines.push(`$rng.NumberFormat = '${escapePowerShellSingleQuoted(fmt.numberFormat)}'`);
    if (fmt.horizontalAlignment) psLines.push(`$rng.HorizontalAlignment = [int][Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlign${fmt.horizontalAlignment}`);
    const psBody = psLines.join("\n    ");
    const script = `
$ErrorActionPreference = "Stop"
try {
    Add-Type -AssemblyName System.Drawing
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(this.filePath.replace(/\//g, "\\\\"))}') }
    $ws = if ($sheetLit -eq "$null") { $wb.Worksheets.Item(1) } else { $wb.Sheets.Item($sheetLit) }
    $rng = $ws.Range('${escapePowerShellSingleQuoted(address)}')
    ${psBody}
    $wb.Save()
    $actualSheet = [string]$ws.Name
    $payload = @{ success = $true; sheetName = $actualSheet; address = '${escapePowerShellSingleQuoted(address)}' }
    Write-Output (ConvertTo-Json $payload -Compress)
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`.replace(/\$sheetLit/g, sheetLit);
    return runPowerShell(script);
  }

  /** 设置 Excel 主窗口置顶或取消置顶（通过 COM 获取当前工作簿实例的 HWND 后调用 Win32 API） */
  async setWindowTopMost(onTop: boolean): Promise<ExcelComResult> {
    const action = onTop ? "置顶" : "取消置顶";
    const wbName = basename(this.filePath);
    const script = `
$ErrorActionPreference = "Stop"
try {
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32TopMost {
        [DllImport(\"user32.dll\", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport(\"user32.dll\", SetLastError = true)]
        public static extern bool IsWindow(IntPtr hWnd);
        [DllImport(\"user32.dll\", SetLastError = true)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport(\"user32.dll\", SetLastError = true)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    }
"@

    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $targetWbName = '${escapePowerShellSingleQuoted(wbName)}'
    $found = $false
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq $targetWbName) { $found = $true; break }
    }
    if (-not $found) { throw "未找到工作簿：$targetWbName" }

    $hwnd = [IntPtr]::new([long]$excel.Hwnd)
    if (-not [Win32TopMost]::IsWindow($hwnd)) { throw "Excel 窗口句柄无效：$hwnd" }

    $HWND_TOPMOST = [IntPtr]::new(-1)
    $HWND_NOTOPMOST = [IntPtr]::new(-2)
    $SWP_NOMOVE = 0x0002
    $SWP_NOSIZE = 0x0001

    # 激活窗口确保置顶命令生效，但不改变窗口的显示状态（最大化/最小化保持原样）
    [void][Win32TopMost]::SetForegroundWindow($hwnd)

    $target = if (${onTop ? "$true" : "$false"}) { $HWND_TOPMOST } else { $HWND_NOTOPMOST }
    $flags = $SWP_NOMOVE -bor $SWP_NOSIZE
    $result = [Win32TopMost]::SetWindowPos($hwnd, $target, 0, 0, 0, 0, $flags)
    if (-not $result) {
        $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "SetWindowPos 调用失败，错误码：$err，HWND：$hwnd"
    }
    Write-Output "Excel 窗口已${action} (HWND=$hwnd)"
} catch {
    $errMsg = [string]$_.Exception.Message
    $errStack = [string]$_.ScriptStackTrace
    Write-Error ("ERROR: " + $errMsg + "\`nSTACK: " + $errStack)
}
`;
    return runPowerShell(script);
  }
}
