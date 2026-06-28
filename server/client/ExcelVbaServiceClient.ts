/**
 * Excel VBA Assistant - MCP 服务客户端
 *
 * 包装 VbaClient，维护 workbook 注册表，对任务文档第二十节的 14 个工具提供方法。
 *
 * 工作簿注册表：
 *   workbookId -> { filePath, localDir, client: VbaClient }
 *
 * 启动时：
 *   - 通过 VBE_FILE_PATH 注册默认工作簿（workbookId = 文件名，不含扩展名）
 *   - 通过 VBE_LOCAL_DIR 提供默认本地同步目录
 *   - 通过 VBE_READ_ONLY=true 启用只读模式，禁止所有危险操作
 */
import { basename, resolve } from "path";

import { VbaClient } from "../../src/client/vbaClient";
import { listExcelDialogs, clickExcelDialog, fillDialogInput } from "../../src/client/vbaMacroRunner";
import { runPowerShell, escapePowerShellSingleQuoted, type ExcelComResult } from "../../src/runtime/powershell";

/** 工作簿注册条目 */
export interface WorkbookRegistryEntry {
  workbookId: string;
  filePath: string;
  localDir?: string;
  client: VbaClient;
}

/** excel_list_workbooks 返回项 */
export interface WorkbookSummary {
  workbookId: string;
  filePath: string;
  localDir?: string;
  isDefault: boolean;
}

export class ExcelVbaServiceClient {
  private readonly registry = new Map<string, WorkbookRegistryEntry>();
  private readonly defaultWorkbookId: string | null;
  private readonly defaultLocalDir: string | null;
  private readonly readOnly: boolean;

  constructor() {
    const envPath = process.env.VBE_FILE_PATH || "";
    this.defaultLocalDir = process.env.VBE_LOCAL_DIR ? resolve(process.env.VBE_LOCAL_DIR) : null;
    this.readOnly = (process.env.VBE_READ_ONLY || "").toLowerCase() === "true";

    if (envPath) {
      const absPath = resolve(envPath);
      const id = this.deriveWorkbookId(absPath);
      const client = new VbaClient(absPath);
      this.registry.set(id, {
        workbookId: id,
        filePath: absPath,
        localDir: this.defaultLocalDir || undefined,
        client,
      });
      this.defaultWorkbookId = id;
    } else {
      this.defaultWorkbookId = null;
    }
  }

  /** 从文件路径推导 workbookId（文件名，不含扩展名） */
  private deriveWorkbookId(filePath: string): string {
    const fileName = basename(filePath);
    const dotIndex = fileName.lastIndexOf(".");
    return dotIndex > 0 ? fileName.slice(0, dotIndex) : fileName;
  }

  isReadOnly(): boolean {
    return this.readOnly;
  }

  getDefaultLocalDir(): string | null {
    return this.defaultLocalDir;
  }

  getDefaultWorkbookId(): string | null {
    return this.defaultWorkbookId;
  }

  /** 注册或更新一个工作簿，返回注册条目 */
  openWorkbook(filePath: string, workbookId?: string): WorkbookRegistryEntry {
    const absPath = resolve(filePath);
    const id = (workbookId && workbookId.trim()) || this.deriveWorkbookId(absPath);
    const existing = this.registry.get(id);
    if (existing && existing.filePath.toLowerCase() === absPath.toLowerCase()) {
      return existing;
    }
    const client = new VbaClient(absPath);
    const entry: WorkbookRegistryEntry = {
      workbookId: id,
      filePath: absPath,
      localDir: this.defaultLocalDir || undefined,
      client,
    };
    this.registry.set(id, entry);
    return entry;
  }

  /** 列出所有已注册工作簿 */
  listWorkbooks(): WorkbookSummary[] {
    const result: WorkbookSummary[] = [];
    for (const entry of this.registry.values()) {
      result.push({
        workbookId: entry.workbookId,
        filePath: entry.filePath,
        localDir: entry.localDir,
        isDefault: entry.workbookId === this.defaultWorkbookId,
      });
    }
    return result;
  }

  /** 取得注册条目；workbookId 为空时回退到默认工作簿；不存在则抛错 */
  getEntry(workbookId?: string): WorkbookRegistryEntry {
    if (workbookId && workbookId.trim()) {
      const entry = this.registry.get(workbookId);
      if (entry) return entry;
    }
    if (this.defaultWorkbookId) {
      const defaultEntry = this.registry.get(this.defaultWorkbookId);
      if (defaultEntry) return defaultEntry;
    }
    if (workbookId && workbookId.trim()) {
      throw new Error(`未找到工作簿 ID: ${workbookId}。请先调用 excel_open_workbook 或检查 VBE_FILE_PATH 环境变量。`);
    }
    throw new Error("未配置默认工作簿。请设置环境变量 VBE_FILE_PATH 或先调用 excel_open_workbook。");
  }

  /** 取得默认工作簿条目；不存在则抛错 */
  getDefaultEntry(): WorkbookRegistryEntry {
    if (!this.defaultWorkbookId) {
      throw new Error("未配置默认工作簿。请设置环境变量 VBE_FILE_PATH 或先调用 excel_open_workbook。");
    }
    return this.getEntry(this.defaultWorkbookId);
  }

  /** 解析 localDir：优先使用参数，其次注册时记录的目录，最后回退到默认 */
  private resolveLocalDir(entry: WorkbookRegistryEntry, explicit?: string): string {
    if (explicit && explicit.trim()) return resolve(explicit);
    if (entry.localDir) return entry.localDir;
    if (this.defaultLocalDir) return this.defaultLocalDir;
    throw new Error("未提供本地同步目录。请通过参数 localDir 或环境变量 VBE_LOCAL_DIR 指定。");
  }

  /** 前置 Excel 运行状态检查（统一包装） */
  private async preCheck(entry: WorkbookRegistryEntry): Promise<ExcelComResult | null> {
    const accessError = await entry.client.checkAccess();
    if (accessError) {
      return { success: false, message: accessError };
    }
    return null;
  }

  // ============================================================
  // 工具方法（对齐任务文档第二十节）
  // ============================================================

  /** excel_list_resources */
  async listResources(workbookId?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.getResources();
  }

  /** excel_inspect_workbook（全景信息） */
  async inspectWorkbook(workbookId?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.inspectWorkbook();
  }

  /** excel_get_all_vba_code */
  async getAllVbaCode(workbookId?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.getAllComponentCode();
  }

  /** excel_get_vba_code */
  async getVbaCode(workbookId: string | undefined, componentName: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.getComponentCode(componentName);
  }

  /** excel_update_vba_code（危险操作） */
  async updateVbaCode(workbookId: string | undefined, componentName: string, code: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止修改 VBA 代码。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    // 直接通过 PowerShell 写入指定组件的代码（覆盖性写入，对应任务文档第十二节）
    const wbName = basename(entry.filePath);
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(entry.filePath.replace(/\//g, "\\"))}') }
    $comp = $null
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Name -eq '${escapePowerShellSingleQuoted(componentName)}') { $comp = $c; break } }
    if ($comp -eq $null) { Write-Error "未找到组件 ${escapePowerShellSingleQuoted(componentName)}"; exit }
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $codeBytes = [System.Convert]::FromBase64String('${Buffer.from(code, "utf-8").toString("base64")}')
    $code = [System.Text.Encoding]::UTF8.GetString($codeBytes)
    if ($code.Length -gt 0) { $cm.AddFromString($code) }
    $wb.Save()
    Write-Output "已更新组件: ${escapePowerShellSingleQuoted(componentName)}"
} catch { Write-Error $_.Exception.Message }
`;
    return runPowerShell(script);
  }

  /** excel_create_vba_component（危险操作） */
  async createVbaComponent(
    workbookId: string | undefined,
    componentType: "standardModule" | "classModule" | "userForm",
    componentName: string,
    code?: string
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止新增 VBA 组件。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;

    const typeCode = componentType === "standardModule" ? 1
      : componentType === "classModule" ? 2
      : componentType === "userForm" ? 3
      : 0;
    if (typeCode === 0) {
      return { success: false, message: `不支持的组件类型: ${componentType}` };
    }

    const wbName = basename(entry.filePath);
    const codeLiteral = code ? Buffer.from(code, "utf-8").toString("base64") : "";
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(entry.filePath.replace(/\//g, "\\"))}') }
    $vbProject = $wb.VBProject
    foreach ($c in $vbProject.VBComponents) {
        if ($c.Name -eq '${escapePowerShellSingleQuoted(componentName)}') { Write-Error "组件 ${escapePowerShellSingleQuoted(componentName)} 已存在"; exit }
    }
    $newComp = $vbProject.VBComponents.Add(${typeCode})
    $newComp.Name = '${escapePowerShellSingleQuoted(componentName)}'
    ${code ? `
    $codeBytes = [System.Convert]::FromBase64String('${codeLiteral}')
    $codeText = [System.Text.Encoding]::UTF8.GetString($codeBytes)
    if ($codeText.Length -gt 0) { $newComp.CodeModule.AddFromString($codeText) }
    ` : ""}
    $wb.Save()
    Write-Output "已创建组件: ${escapePowerShellSingleQuoted(componentName)} (type=${typeCode})"
} catch { Write-Error $_.Exception.Message }
`;
    return runPowerShell(script);
  }

  /** excel_delete_vba_component（危险操作） */
  async deleteVbaComponent(workbookId: string | undefined, componentName: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止删除 VBA 组件。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;

    const wbName = basename(entry.filePath);
    const script = `
$ErrorActionPreference = "Stop"
try {
    $excel = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $null
    foreach ($w in $excel.Workbooks) { if ($w.Name -eq '${escapePowerShellSingleQuoted(wbName)}') { $wb = $w; break } }
    if ($wb -eq $null) { $wb = $excel.Workbooks.Open('${escapePowerShellSingleQuoted(entry.filePath.replace(/\//g, "\\"))}') }
    $vbProject = $wb.VBProject
    $target = $null
    foreach ($c in $vbProject.VBComponents) {
        if ($c.Name -eq '${escapePowerShellSingleQuoted(componentName)}') { $target = $c; break }
    }
    if ($target -eq $null) { Write-Error "未找到组件 ${escapePowerShellSingleQuoted(componentName)}"; exit }
    $typeCode = [int]$target.Type
    if ($typeCode -eq 100) { Write-Error "工作表/文档对象模块不允许删除，请在 Excel 中删除对应工作表"; exit }
    $vbProject.VBComponents.Remove($target)
    $wb.Save()
    Write-Output "已删除组件: ${escapePowerShellSingleQuoted(componentName)}"
} catch { Write-Error $_.Exception.Message }
`;
    return runPowerShell(script);
  }

  /** excel_list_macros */
  async listMacros(workbookId?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.listMacros();
  }

  /** excel_run_macro（危险操作） */
  async runMacro(
    workbookId: string | undefined,
    macroName: string,
    args?: unknown[],
    timeoutSeconds?: number,
    captureResultRange?: string
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止运行宏。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    // 第一版仅支持无参宏；若调用方传入 args，给出明确提示
    if (args && args.length > 0) {
      return { success: false, message: "当前版本仅支持无参数宏执行。请在 VBA 中改用模块级状态变量传参。" };
    }
    return entry.client.runMacro(macroName, {
      timeoutMs: timeoutSeconds ? timeoutSeconds * 1000 : undefined,
      captureResultRange,
    });
  }

  /** excel_list_dialogs（不依赖特定工作簿） */
  async listDialogs(): Promise<ExcelComResult> {
    return listExcelDialogs();
  }

  /** excel_click_dialog（危险操作，不依赖特定工作簿） */
  async clickDialog(handle: string, action?: string, buttonText?: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止点击弹窗。" };
    }
    return clickExcelDialog(handle, action, buttonText);
  }

  /** excel_fill_dialog（危险操作，不依赖特定工作簿） */
  async fillDialog(handle: string, text: string, submit = false): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止向弹窗输入文本。" };
    }
    return fillDialogInput(handle, text, submit);
  }

  /** excel_list_sheets */
  async listSheets(workbookId?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.listSheets();
  }

  /** excel_read_range */
  async readRange(workbookId: string | undefined, sheetName: string | undefined, address: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.readRange(sheetName, address);
  }

  /** excel_read_used_range */
  async readUsedRange(workbookId: string | undefined, sheetName: string | undefined): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.readUsedRange(sheetName);
  }

  /** excel_sync_vbe_to_local */
  async syncVbeToLocal(workbookId: string | undefined, localDir?: string): Promise<ExcelComResult> {
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    const dir = this.resolveLocalDir(entry, localDir);
    const result = await entry.client.syncVbeToLocal(dir);
    // 同步成功后记录到注册表，便于后续工具使用
    if (result.success) {
      entry.localDir = dir;
    }
    return result;
  }

  /** excel_sync_local_to_vbe（危险操作：覆盖性同步） */
  async syncLocalToVbe(workbookId: string | undefined, localDir?: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止 本地 → VBE 同步。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    const dir = this.resolveLocalDir(entry, localDir);
    const result = await entry.client.syncLocalToVbe(dir);
    if (result.success) {
      entry.localDir = dir;
    }
    return result;
  }

  // ============================================================
  // Excel 工作表 / 单元格 编辑（需求 1）
  // ============================================================

  /** excel_create_sheet */
  async createSheet(
    workbookId: string | undefined,
    sheetName: string,
    before?: string,
    after?: string
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止创建工作表。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.createSheet(sheetName, { before, after });
  }

  /** excel_delete_sheet */
  async deleteSheet(workbookId: string | undefined, sheetName: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止删除工作表。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.deleteSheet(sheetName);
  }

  /** excel_rename_sheet */
  async renameSheet(workbookId: string | undefined, oldName: string, newName: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止重命名工作表。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.renameSheet(oldName, newName);
  }

  /** excel_set_cell_value */
  async setCellValue(
    workbookId: string | undefined,
    sheetName: string | undefined,
    address: string,
    value: unknown
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止写入单元格。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.setCellValue(sheetName, address, value as import("../../src/client/vbaClient").CellValue);
  }

  /** excel_set_range_values */
  async setRangeValues(
    workbookId: string | undefined,
    sheetName: string | undefined,
    startAddress: string,
    values: unknown[][]
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止批量写入单元格。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.setRangeValues(sheetName, startAddress, values as import("../../src/client/vbaClient").CellValue[][]);
  }

  /** excel_clear_range */
  async clearRange(workbookId: string | undefined, sheetName: string | undefined, address: string): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止清空单元格。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.clearRange(sheetName, address);
  }

  /** excel_set_cell_format */
  async setCellFormat(
    workbookId: string | undefined,
    sheetName: string | undefined,
    address: string,
    format: Record<string, unknown>
  ): Promise<ExcelComResult> {
    if (this.readOnly) {
      return { success: false, message: "只读模式已启用（VBE_READ_ONLY=true），禁止设置单元格格式。" };
    }
    const entry = this.getEntry(workbookId);
    const pre = await this.preCheck(entry);
    if (pre) return pre;
    return entry.client.setCellFormat(sheetName, address, format as import("../../src/client/vbaClient").CellFormat);
  }
}
