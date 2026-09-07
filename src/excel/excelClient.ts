/**
 * Excel Client
 *
 * 替代原 VbaClient，所有 Excel 操作均通过 mcp-server-excel MCP 工具完成。
 * 封装 mcp-server-excel 全部 MCP 工具；常用操作提供类型化方法，其余可通过 callTool 直接调用。
 */
import { createHash } from "crypto";
import { McpClientWrapper, extractJson } from "../mcp/mcpClient";
import { SessionStore } from "./sessionStore";

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  numberFormat?: string;
  horizontalAlignment?: "Left" | "Center" | "Right";
}

export type CellValue = string | number | boolean | null;

interface SnakeCellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font_color?: string;
  fill_color?: string;
  font_size?: number;
  number_format?: string;
  horizontal_alignment?: string;
}

export class ExcelClient {
  private sessionStore: SessionStore;

  constructor(private client: McpClientWrapper) {
    this.sessionStore = new SessionStore(client);
  }

  getSessionStore(): SessionStore {
    return this.sessionStore;
  }

  // ============================================================
  // Low-level tool call
  // ============================================================

  async callTool(toolName: string, args: Record<string, unknown>): Promise<unknown> {
    const result = await this.client.callTool({ name: toolName, arguments: args });
    return extractJson(result);
  }

  // ============================================================
  // Session / File
  // ============================================================

  async openWorkbook(filePath: string): Promise<string> {
    // T1: 用户选择文件时通过 MCP 打开 Excel，并确保窗口可见
    return this.sessionStore.ensureSession(filePath, { show: true });
  }

  async closeWorkbook(filePath: string, save = true): Promise<void> {
    return this.sessionStore.closeSession(filePath, save);
  }

  async listSessions(): Promise<unknown> {
    return this.callTool("file", { action: "list" });
  }

  async createEmptyWorkbook(filePath: string): Promise<unknown> {
    return this.callTool("file", { action: "create", path: filePath });
  }

  async testWorkbook(filePath: string): Promise<unknown> {
    return this.callTool("file", { action: "test", path: filePath });
  }

  async closeWorkbookOnly(filePath: string, save = true): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(filePath);
    await this.callTool("file", { action: "close-workbook", session_id, save });
  }

  // ============================================================
  // VBA
  // ============================================================

  async vbaList(workbookPath: string): Promise<Array<{ name: string; type: string }>> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("vba", { session_id, action: "list" });
    return (data as { modules?: Array<{ name: string; type: string }> })?.modules ?? [];
  }

  async vbaView(workbookPath: string, moduleName: string): Promise<string> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("vba", { session_id, action: "view", module_name: moduleName });
    return (data as { code?: string })?.code ?? "";
  }

  async vbaImport(workbookPath: string, moduleName: string, code: string, type: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("vba", { session_id, action: "import", module_name: moduleName, vba_code: code, type });
  }

  async vbaUpdate(workbookPath: string, moduleName: string, code: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("vba", { session_id, action: "update", module_name: moduleName, vba_code: code });
  }

  async vbaDelete(workbookPath: string, moduleName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("vba", { session_id, action: "delete", module_name: moduleName });
  }

  async vbaRun(
    workbookPath: string,
    macroName: string,
    parameters?: string[] | string
  ): Promise<{ success: boolean; message: string }> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const params: string[] | undefined =
      parameters === undefined
        ? undefined
        : Array.isArray(parameters)
          ? parameters
          : parameters
            ? [parameters]
            : undefined;
    const data = await this.callTool("vba", {
      session_id,
      action: "run",
      procedure_name: macroName,
      parameters: params,
    });
    return {
      success: (data as { success?: boolean })?.success ?? true,
      message: (data as { message?: string })?.message ?? "宏执行完成",
    };
  }

  // ============================================================
  // Worksheets
  // ============================================================

  async listSheets(workbookPath: string): Promise<Array<{ name: string; index: number }>> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("worksheet", { session_id, action: "list" });
    return (data as { sheets?: Array<{ name: string; index: number }> })?.sheets ?? [];
  }

  async createSheet(
    workbookPath: string,
    name: string,
    options?: { before?: string; after?: string }
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet", {
      session_id,
      action: "create",
      sheet_name: name,
      before_sheet: options?.before,
      after_sheet: options?.after,
    });
  }

  async deleteSheet(workbookPath: string, name: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet", { session_id, action: "delete", sheet_name: name });
  }

  async renameSheet(workbookPath: string, oldName: string, newName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet", { session_id, action: "rename", old_name: oldName, new_name: newName });
  }

  async copySheet(workbookPath: string, sourceName: string, targetName?: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet", {
      session_id,
      action: "copy",
      source_name: sourceName,
      target_name: targetName,
    });
  }

  async moveSheet(
    workbookPath: string,
    sheetName: string,
    options?: { beforeSheet?: string; afterSheet?: string }
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet", {
      session_id,
      action: "move",
      sheet_name: sheetName,
      before_sheet: options?.beforeSheet,
      after_sheet: options?.afterSheet,
    });
  }

  async copySheetToFile(sourceFile: string, sourceSheet: string, targetFile: string, targetSheetName?: string): Promise<void> {
    await this.callTool("worksheet", {
      action: "copy-to-file",
      source_file: sourceFile,
      source_sheet: sourceSheet,
      target_file: targetFile,
      target_sheet_name: targetSheetName,
    });
  }

  async moveSheetToFile(sourceFile: string, sourceSheet: string, targetFile: string, targetSheetName?: string): Promise<void> {
    await this.callTool("worksheet", {
      action: "move-to-file",
      source_file: sourceFile,
      source_sheet: sourceSheet,
      target_file: targetFile,
      target_sheet_name: targetSheetName,
    });
  }

  async setSheetTabColor(workbookPath: string, name: string, color?: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const rgb = this.parseColor(color);
    await this.callTool("worksheet_style", {
      session_id,
      action: "set-tab-color",
      sheet_name: name,
      red: rgb.red,
      green: rgb.green,
      blue: rgb.blue,
    });
  }

  async setSheetVisibility(workbookPath: string, name: string, visibility: "visible" | "hidden" | "very-hidden"): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("worksheet_style", {
      session_id,
      action: "set-visibility",
      sheet_name: name,
      visibility,
    });
  }

  // ============================================================
  // Ranges
  // ============================================================

  async readRange(workbookPath: string, sheetName: string | undefined, address: string): Promise<CellValue[][]> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("range", {
      session_id,
      action: "get-values",
      sheet_name: sheetName,
      range_address: address,
    });
    return (data as { values?: CellValue[][] })?.values ?? [];
  }

  async getRangeFormulas(workbookPath: string, sheetName: string | undefined, address: string): Promise<CellValue[][]> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("range", {
      session_id,
      action: "get-formulas",
      sheet_name: sheetName,
      range_address: address,
    });
    return (data as { formulas?: CellValue[][] })?.formulas ?? [];
  }

  async setRangeValues(
    workbookPath: string,
    sheetName: string | undefined,
    startAddress: string,
    values: CellValue[][]
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range", {
      session_id,
      action: "set-values",
      sheet_name: sheetName,
      range_address: startAddress,
      values,
    });
  }

  async setRangeFormulas(
    workbookPath: string,
    sheetName: string | undefined,
    startAddress: string,
    formulas: CellValue[][]
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range", {
      session_id,
      action: "set-formulas",
      sheet_name: sheetName,
      range_address: startAddress,
      formulas,
    });
  }

  async clearRange(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    what: "all" | "contents" | "formats" = "contents"
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const action = what === "all" ? "clear-all" : what === "formats" ? "clear-formats" : "clear-contents";
    await this.callTool("range", { session_id, action, sheet_name: sheetName, range_address: address });
  }

  async getUsedRange(workbookPath: string, sheetName?: string): Promise<CellValue[][]> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("range", { session_id, action: "get-used-range", sheet_name: sheetName });
    return (data as { values?: CellValue[][] })?.values ?? [];
  }

  async getCurrentRegion(workbookPath: string, sheetName: string, cellAddress: string): Promise<CellValue[][]> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const data = await this.callTool("range", {
      session_id,
      action: "get-current-region",
      sheet_name: sheetName,
      cell_address: cellAddress,
    });
    return (data as { values?: CellValue[][] })?.values ?? [];
  }

  async copyRange(
    workbookPath: string,
    sourceSheet: string,
    sourceRange: string,
    targetSheet: string,
    targetRange: string,
    what?: "all" | "values" | "formulas"
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    const action = what === "values" ? "copy-values" : what === "formulas" ? "copy-formulas" : "copy";
    await this.callTool("range", {
      session_id,
      action,
      source_sheet: sourceSheet,
      source_range: sourceRange,
      target_sheet: targetSheet,
      target_range: targetRange,
    });
  }

  async setNumberFormat(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    numberFormat: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range", {
      session_id,
      action: "set-number-format",
      sheet_name: sheetName,
      range_address: address,
      format_code: numberFormat,
    });
  }

  // ============================================================
  // Range Edit (insert/delete/find/sort)
  // ============================================================

  async sortRange(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    sortColumns: Array<{ columnIndex: number; ascending: boolean }>
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_edit", {
      session_id,
      action: "sort",
      sheet_name: sheetName,
      range_address: address,
      sort_columns: sortColumns,
    });
  }

  async findInRange(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    searchValue: string,
    options?: { matchCase?: boolean; matchEntireCell?: boolean }
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("range_edit", {
      session_id,
      action: "find",
      sheet_name: sheetName,
      range_address: address,
      search_value: searchValue,
      find_options: options,
    });
  }

  async replaceInRange(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    findValue: string,
    replaceValue: string,
    replaceAll = true
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("range_edit", {
      session_id,
      action: "replace",
      sheet_name: sheetName,
      range_address: address,
      find_value: findValue,
      replace_value: replaceValue,
      replace_options: { replaceAll },
    });
  }

  async insertRows(workbookPath: string, sheetName: string, row: number, count = 1): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_edit", {
      session_id,
      action: "insert-rows",
      sheet_name: sheetName,
      range_address: `${row}:${row + count - 1}`,
    });
  }

  async deleteRows(workbookPath: string, sheetName: string, row: number, count = 1): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_edit", {
      session_id,
      action: "delete-rows",
      sheet_name: sheetName,
      range_address: `${row}:${row + count - 1}`,
    });
  }

  async insertColumns(workbookPath: string, sheetName: string, column: string, count = 1): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_edit", {
      session_id,
      action: "insert-columns",
      sheet_name: sheetName,
      range_address: `${column}:${this.shiftColumn(column, count - 1)}`,
    });
  }

  async deleteColumns(workbookPath: string, sheetName: string, column: string, count = 1): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_edit", {
      session_id,
      action: "delete-columns",
      sheet_name: sheetName,
      range_address: `${column}:${this.shiftColumn(column, count - 1)}`,
    });
  }

  // ============================================================
  // Range Format
  // ============================================================

  async setCellFormat(workbookPath: string, sheetName: string | undefined, address: string, format: CellFormat): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "format-range",
      sheet_name: sheetName,
      range_address: address,
      ...this.toSnakeFormat(format),
    });
  }

  async formatRanges(
    workbookPath: string,
    sheetName: string,
    rangeAddresses: string[],
    format: CellFormat
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "format-ranges",
      sheet_name: sheetName,
      range_addresses: rangeAddresses,
      ...this.toSnakeFormat(format),
    });
  }

  async setRangeStyle(workbookPath: string, sheetName: string | undefined, address: string, styleName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "set-style",
      sheet_name: sheetName,
      range_address: address,
      style_name: styleName,
    });
  }

  async autoFitColumns(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "auto-fit-columns",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async autoFitRows(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "auto-fit-rows",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async mergeCells(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "merge-cells",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async unmergeCells(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "unmerge-cells",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async addDataValidation(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    validation: Record<string, unknown>
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "validate-range",
      sheet_name: sheetName,
      range_address: address,
      ...validation,
    });
  }

  async removeDataValidation(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "remove-validation",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async setColumnWidth(workbookPath: string, sheetName: string | undefined, address: string, width: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "set-column-width",
      sheet_name: sheetName,
      range_address: address,
      column_width: width,
    });
  }

  async setRowHeight(workbookPath: string, sheetName: string | undefined, address: string, height: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_format", {
      session_id,
      action: "set-row-height",
      sheet_name: sheetName,
      range_address: address,
      row_height: height,
    });
  }

  // ============================================================
  // Range Link (hyperlinks / cell protection)
  // ============================================================

  async addHyperlink(
    workbookPath: string,
    sheetName: string,
    cellAddress: string,
    url: string,
    displayText?: string,
    tooltip?: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_link", {
      session_id,
      action: "add-hyperlink",
      sheet_name: sheetName,
      cell_address: cellAddress,
      url,
      display_text: displayText,
      tooltip,
    });
  }

  async removeHyperlink(workbookPath: string, sheetName: string, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_link", {
      session_id,
      action: "remove-hyperlink",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  async listHyperlinks(workbookPath: string, sheetName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("range_link", { session_id, action: "list-hyperlinks", sheet_name: sheetName });
  }

  async setCellLock(workbookPath: string, sheetName: string, address: string, locked: boolean): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("range_link", {
      session_id,
      action: "set-cell-lock",
      sheet_name: sheetName,
      range_address: address,
      locked,
    });
  }

  // ============================================================
  // Calculation Mode
  // ============================================================

  async getCalculationMode(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("calculation_mode", { session_id, action: "get-mode" });
  }

  async setCalculationMode(workbookPath: string, mode: "automatic" | "manual" | "semi-automatic"): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("calculation_mode", { session_id, action: "set-mode", mode });
  }

  async calculate(
    workbookPath: string,
    scope: "workbook" | "sheet" | "range" = "workbook",
    sheetName?: string,
    rangeAddress?: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("calculation_mode", {
      session_id,
      action: "calculate",
      scope,
      sheet_name: sheetName,
      range_address: rangeAddress,
    });
  }

  // ============================================================
  // Excel Tables
  // ============================================================

  async listTables(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", { session_id, action: "list" });
  }

  async createTable(
    workbookPath: string,
    sheetName: string,
    rangeAddress: string,
    tableName?: string,
    options?: { hasHeaders?: boolean; styleName?: string }
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", {
      session_id,
      action: "create",
      sheet_name: sheetName,
      range_address: rangeAddress,
      table_name: tableName,
      has_headers: options?.hasHeaders ?? true,
      table_style: options?.styleName,
    });
  }

  async renameTable(workbookPath: string, tableName: string, newName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", { session_id, action: "rename", table_name: tableName, new_name: newName });
  }

  async resizeTable(workbookPath: string, tableName: string, newRange: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", { session_id, action: "resize", table_name: tableName, new_range: newRange });
  }

  async getTableData(workbookPath: string, tableName: string, visibleOnly = false): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", { session_id, action: "get-data", table_name: tableName, visible_only: visibleOnly });
  }

  async appendTableRows(workbookPath: string, tableName: string, rows: CellValue[][]): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("table", { session_id, action: "append", table_name: tableName, rows });
  }

  async toggleTableTotals(workbookPath: string, tableName: string, showTotals: boolean): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", {
      session_id,
      action: "toggle-totals",
      table_name: tableName,
      show_totals: showTotals,
    });
  }

  async setTableColumnTotal(
    workbookPath: string,
    tableName: string,
    columnName: string,
    totalFunction: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", {
      session_id,
      action: "set-column-total",
      table_name: tableName,
      column_name: columnName,
      total_function: totalFunction,
    });
  }

  async setTableStyle(workbookPath: string, tableName: string, styleName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table", { session_id, action: "set-style", table_name: tableName, table_style: styleName });
  }

  async addTableToDataModel(workbookPath: string, tableName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("table", { session_id, action: "add-to-data-model", table_name: tableName });
  }

  async deleteTable(workbookPath: string, tableName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("table", { session_id, action: "delete", table_name: tableName });
  }

  // ============================================================
  // Table Column
  // ============================================================

  async applyTableFilter(workbookPath: string, tableName: string, columnName: string, criteria: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "apply-filter",
      table_name: tableName,
      column_name: columnName,
      criteria,
    });
  }

  async applyTableFilterValues(workbookPath: string, tableName: string, columnName: string, values: string[]): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "apply-filter-values",
      table_name: tableName,
      column_name: columnName,
      values: JSON.stringify(values),
    });
  }

  async clearTableFilters(workbookPath: string, tableName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", { session_id, action: "clear-filters", table_name: tableName });
  }

  async sortTableColumn(
    workbookPath: string,
    tableName: string,
    columnName: string,
    ascending = true
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "sort",
      table_name: tableName,
      column_name: columnName,
      ascending,
    });
  }

  async addTableColumn(
    workbookPath: string,
    tableName: string,
    columnName: string,
    position?: number
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "add-column",
      table_name: tableName,
      column_name: columnName,
      position,
    });
  }

  async removeTableColumn(workbookPath: string, tableName: string, columnName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "remove-column",
      table_name: tableName,
      column_name: columnName,
    });
  }

  async renameTableColumn(
    workbookPath: string,
    tableName: string,
    oldName: string,
    newName: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("table_column", {
      session_id,
      action: "rename-column",
      table_name: tableName,
      old_name: oldName,
      new_name: newName,
    });
  }

  // ============================================================
  // PivotTables
  // ============================================================

  async listPivotTables(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable", { session_id, action: "list" });
  }

  async readPivotTable(workbookPath: string, pivotTableName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable", { session_id, action: "read", pivot_table_name: pivotTableName });
  }

  async createPivotTableFromRange(
    workbookPath: string,
    pivotTableName: string,
    sourceSheet: string,
    sourceRange: string,
    destinationSheet: string,
    destinationCell: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable", {
      session_id,
      action: "create-from-range",
      pivot_table_name: pivotTableName,
      source_sheet: sourceSheet,
      source_range: sourceRange,
      destination_sheet: destinationSheet,
      destination_cell: destinationCell,
    });
  }

  async createPivotTableFromTable(
    workbookPath: string,
    pivotTableName: string,
    tableName: string,
    destinationSheet: string,
    destinationCell: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable", {
      session_id,
      action: "create-from-table",
      pivot_table_name: pivotTableName,
      table_name: tableName,
      destination_sheet: destinationSheet,
      destination_cell: destinationCell,
    });
  }

  async refreshPivotTable(workbookPath: string, pivotTableName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable", { session_id, action: "refresh", pivot_table_name: pivotTableName });
  }

  async deletePivotTable(workbookPath: string, pivotTableName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable", { session_id, action: "delete", pivot_table_name: pivotTableName });
  }

  // ============================================================
  // PivotTable Field
  // ============================================================

  async addPivotRowField(workbookPath: string, pivotTableName: string, fieldName: string, position?: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "add-row-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      position,
    });
  }

  async addPivotColumnField(workbookPath: string, pivotTableName: string, fieldName: string, position?: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "add-column-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      position,
    });
  }

  async addPivotValueField(workbookPath: string, pivotTableName: string, fieldName: string, position?: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "add-value-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      position,
    });
  }

  async addPivotFilterField(workbookPath: string, pivotTableName: string, fieldName: string, position?: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "add-filter-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      position,
    });
  }

  async removePivotField(workbookPath: string, pivotTableName: string, fieldName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "remove-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
    });
  }

  async setPivotFieldFunction(
    workbookPath: string,
    pivotTableName: string,
    fieldName: string,
    aggregationFunction: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "set-field-function",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      aggregation_function: aggregationFunction,
    });
  }

  async setPivotFieldName(
    workbookPath: string,
    pivotTableName: string,
    fieldName: string,
    customName: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "set-field-name",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      custom_name: customName,
    });
  }

  async setPivotFieldFilter(
    workbookPath: string,
    pivotTableName: string,
    fieldName: string,
    selectedValues: string[]
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("pivottable_field", {
      session_id,
      action: "set-field-filter",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      selected_values: JSON.stringify(selectedValues),
    });
  }

  // ============================================================
  // PivotTable Calc
  // ============================================================

  async getPivotTableData(workbookPath: string, pivotTableName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable_calc", { session_id, action: "get-data", pivot_table_name: pivotTableName });
  }

  async createPivotCalculatedField(
    workbookPath: string,
    pivotTableName: string,
    fieldName: string,
    formula: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable_calc", {
      session_id,
      action: "create-calculated-field",
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      formula,
    });
  }

  async setPivotTableLayout(
    workbookPath: string,
    pivotTableName: string,
    rowLayout: 0 | 1 | 2
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable_calc", {
      session_id,
      action: "set-layout",
      pivot_table_name: pivotTableName,
      row_layout: rowLayout,
    });
  }

  async setPivotTableGrandTotals(
    workbookPath: string,
    pivotTableName: string,
    showRowGrandTotals: boolean,
    showColumnGrandTotals: boolean
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("pivottable_calc", {
      session_id,
      action: "set-grand-totals",
      pivot_table_name: pivotTableName,
      show_row_grand_totals: showRowGrandTotals,
      show_column_grand_totals: showColumnGrandTotals,
    });
  }

  // ============================================================
  // Charts
  // ============================================================

  async listCharts(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart", { session_id, action: "list" });
  }

  async readChart(workbookPath: string, chartName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart", { session_id, action: "read", chart_name: chartName });
  }

  async createChartFromRange(
    workbookPath: string,
    chartName: string,
    sheetName: string,
    sourceRangeAddress: string,
    chartType: string,
    options?: { targetRange?: string; left?: number; top?: number; width?: number; height?: number }
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart", {
      session_id,
      action: "create-from-range",
      chart_name: chartName,
      sheet_name: sheetName,
      source_range_address: sourceRangeAddress,
      chart_type: chartType,
      target_range: options?.targetRange,
      left: options?.left,
      top: options?.top,
      width: options?.width,
      height: options?.height,
    });
  }

  async deleteChart(workbookPath: string, chartName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("chart", { session_id, action: "delete", chart_name: chartName });
  }

  async moveChart(workbookPath: string, chartName: string, targetRange?: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart", { session_id, action: "move", chart_name: chartName, target_range: targetRange });
  }

  // ============================================================
  // Chart Config
  // ============================================================

  async setChartTitle(workbookPath: string, chartName: string, title: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("chart_config", { session_id, action: "set-title", chart_name: chartName, title });
  }

  async setChartSourceRange(workbookPath: string, chartName: string, sourceRange: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart_config", {
      session_id,
      action: "set-source-range",
      chart_name: chartName,
      source_range: sourceRange,
    });
  }

  async setChartType(workbookPath: string, chartName: string, chartType: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart_config", {
      session_id,
      action: "set-chart-type",
      chart_name: chartName,
      chart_type: chartType,
    });
  }

  async showChartLegend(workbookPath: string, chartName: string, visible: boolean, position?: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("chart_config", {
      session_id,
      action: "show-legend",
      chart_name: chartName,
      visible,
      legend_position: position,
    });
  }

  // ============================================================
  // Power Query
  // ============================================================

  async listQueries(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", { session_id, action: "list" });
  }

  async viewQuery(workbookPath: string, queryName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", { session_id, action: "view", query_name: queryName });
  }

  async evaluateQuery(workbookPath: string, mCode: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", { session_id, action: "evaluate", m_code: mCode });
  }

  async createQuery(workbookPath: string, queryName: string, mCode: string, loadDestination?: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", {
      session_id,
      action: "create",
      query_name: queryName,
      m_code: mCode,
      load_destination: loadDestination,
    });
  }

  async updateQuery(workbookPath: string, queryName: string, mCode: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", { session_id, action: "update", query_name: queryName, m_code: mCode });
  }

  async refreshQuery(workbookPath: string, queryName: string, timeoutSeconds?: number): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("powerquery", {
      session_id,
      action: "refresh",
      query_name: queryName,
      timeout_seconds: timeoutSeconds,
    });
  }

  async refreshAllQueries(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("powerquery", { session_id, action: "refresh-all" });
  }

  async renameQuery(workbookPath: string, queryName: string, newName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", {
      session_id,
      action: "rename",
      old_name: queryName,
      new_name: newName,
    });
  }

  async loadQueryTo(
    workbookPath: string,
    queryName: string,
    loadDestination?: string,
    targetSheet?: string,
    targetCellAddress?: string,
    refresh = true
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("powerquery", {
      session_id,
      action: "load-to",
      query_name: queryName,
      load_destination: loadDestination,
      target_sheet: targetSheet,
      target_cell_address: targetCellAddress,
      refresh,
    });
  }

  async unloadQuery(workbookPath: string, queryName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("powerquery", { session_id, action: "unload", query_name: queryName });
  }

  async deleteQuery(workbookPath: string, queryName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("powerquery", { session_id, action: "delete", query_name: queryName });
  }

  // ============================================================
  // Data Model / DAX
  // ============================================================

  async listDataModelTables(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "list-tables" });
  }

  async readDataModelTable(workbookPath: string, tableName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "read-table", table_name: tableName });
  }

  async renameDataModelTable(workbookPath: string, oldName: string, newName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", {
      session_id,
      action: "rename-table",
      old_name: oldName,
      new_name: newName,
    });
  }

  async deleteDataModelTable(workbookPath: string, tableName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("datamodel", { session_id, action: "delete-table", table_name: tableName });
  }

  async listDataModelColumns(workbookPath: string, tableName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "list-columns", table_name: tableName });
  }

  async readDataModelInfo(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "read-info" });
  }

  async listMeasures(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "list-measures" });
  }

  async readMeasure(workbookPath: string, measureName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "read", measure_name: measureName });
  }

  async createMeasure(
    workbookPath: string,
    tableName: string,
    measureName: string,
    dax: string,
    formatType?: string,
    description?: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", {
      session_id,
      action: "create-measure",
      table_name: tableName,
      measure_name: measureName,
      dax_formula: dax,
      format_type: formatType,
      description,
    });
  }

  async updateMeasure(
    workbookPath: string,
    measureName: string,
    dax: string,
    formatType?: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", {
      session_id,
      action: "update-measure",
      measure_name: measureName,
      dax_formula: dax,
      format_type: formatType,
    });
  }

  async deleteMeasure(workbookPath: string, measureName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("datamodel", { session_id, action: "delete-measure", measure_name: measureName });
  }

  async evaluateDax(workbookPath: string, daxQuery: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "evaluate", dax_query: daxQuery });
  }

  async executeDmv(workbookPath: string, dmvQuery: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel", { session_id, action: "execute-dmv", dmv_query: dmvQuery });
  }

  async refreshDataModel(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("datamodel", { session_id, action: "refresh" });
  }

  // ============================================================
  // Data Model Relationships
  // ============================================================

  async listRelationships(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel_relationship", { session_id, action: "list-relationships" });
  }

  async createRelationship(
    workbookPath: string,
    fromTable: string,
    fromColumn: string,
    toTable: string,
    toColumn: string,
    active = true
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel_relationship", {
      session_id,
      action: "create-relationship",
      from_table: fromTable,
      from_column: fromColumn,
      to_table: toTable,
      to_column: toColumn,
      active,
    });
  }

  async updateRelationship(
    workbookPath: string,
    fromTable: string,
    fromColumn: string,
    toTable: string,
    toColumn: string,
    active: boolean
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("datamodel_relationship", {
      session_id,
      action: "update-relationship",
      from_table: fromTable,
      from_column: fromColumn,
      to_table: toTable,
      to_column: toColumn,
      active,
    });
  }

  async deleteRelationship(
    workbookPath: string,
    fromTable: string,
    fromColumn: string,
    toTable: string,
    toColumn: string
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("datamodel_relationship", {
      session_id,
      action: "delete-relationship",
      from_table: fromTable,
      from_column: fromColumn,
      to_table: toTable,
      to_column: toColumn,
    });
  }

  // ============================================================
  // Named Ranges
  // ============================================================

  async listNamedRanges(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("namedrange", { session_id, action: "list" });
  }

  async readNamedRange(workbookPath: string, name: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("namedrange", { session_id, action: "read", name });
  }

  async writeNamedRange(workbookPath: string, name: string, value: unknown): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("namedrange", { session_id, action: "write", name, value });
  }

  async createNamedRange(workbookPath: string, name: string, refersTo: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("namedrange", { session_id, action: "create", name, reference: refersTo });
  }

  async updateNamedRange(workbookPath: string, name: string, refersTo: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("namedrange", { session_id, action: "update", name, reference: refersTo });
  }

  async deleteNamedRange(workbookPath: string, name: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("namedrange", { session_id, action: "delete", name });
  }

  // ============================================================
  // Data Connections
  // ============================================================

  async listConnections(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "list" });
  }

  async viewConnection(workbookPath: string, connectionName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "view", connection_name: connectionName });
  }

  async createConnection(workbookPath: string, options: Record<string, unknown>): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "create", ...options });
  }

  async testConnection(workbookPath: string, connectionName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "test", connection_name: connectionName });
  }

  async refreshConnection(workbookPath: string, connectionName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("connection", { session_id, action: "refresh", connection_name: connectionName });
  }

  async deleteConnection(workbookPath: string, connectionName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("connection", { session_id, action: "delete", connection_name: connectionName });
  }

  async loadConnectionTo(
    workbookPath: string,
    connectionName: string,
    sheetName: string,
    rangeAddress: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", {
      session_id,
      action: "load-to",
      connection_name: connectionName,
      sheet_name: sheetName,
      range_address: rangeAddress,
    });
  }

  async getConnectionProperties(workbookPath: string, connectionName: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "get-properties", connection_name: connectionName });
  }

  async setConnectionProperties(
    workbookPath: string,
    connectionName: string,
    properties: Record<string, unknown>
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("connection", { session_id, action: "set-properties", connection_name: connectionName, ...properties });
  }

  // ============================================================
  // Slicers
  // ============================================================

  async listSlicers(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("slicer", { session_id, action: "list-slicers" });
  }

  async createSlicer(
    workbookPath: string,
    slicerName: string,
    pivotTableName: string,
    fieldName: string,
    destinationSheet: string,
    position: string
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("slicer", {
      session_id,
      action: "create-slicer",
      slicer_name: slicerName,
      pivot_table_name: pivotTableName,
      field_name: fieldName,
      destination_sheet: destinationSheet,
      position,
    });
  }

  async setSlicerSelection(workbookPath: string, slicerName: string, selectedItems: string[], clearFirst = true): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("slicer", {
      session_id,
      action: "set-slicer-selection",
      slicer_name: slicerName,
      selected_items: JSON.stringify(selectedItems),
      clear_first: clearFirst,
    });
  }

  async deleteSlicer(workbookPath: string, slicerName: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("slicer", { session_id, action: "delete-slicer", slicer_name: slicerName });
  }

  // ============================================================
  // Conditional Formatting
  // ============================================================

  async addConditionalFormat(
    workbookPath: string,
    sheetName: string | undefined,
    address: string,
    rule: Record<string, unknown>
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("conditionalformat", {
      session_id,
      action: "add-rule",
      sheet_name: sheetName,
      range_address: address,
      ...rule,
    });
  }

  async clearConditionalFormat(workbookPath: string, sheetName: string | undefined, address: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("conditionalformat", {
      session_id,
      action: "clear-rules",
      sheet_name: sheetName,
      range_address: address,
    });
  }

  // ============================================================
  // Screenshot
  // ============================================================

  async captureRange(
    workbookPath: string,
    sheetName: string,
    rangeAddress: string,
    quality: "Medium" | "High" | "Low" = "Medium"
  ): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("screenshot", {
      session_id,
      action: "capture",
      sheet_name: sheetName,
      range_address: rangeAddress,
      quality,
    });
  }

  async captureSheet(workbookPath: string, sheetName?: string, quality: "Medium" | "High" | "Low" = "Medium"): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("screenshot", {
      session_id,
      action: "capture-sheet",
      sheet_name: sheetName,
      quality,
    });
  }

  // ============================================================
  // Window Management
  // ============================================================

  async setWindowTopMost(workbookPath: string, onTop: boolean): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: onTop ? "set-topmost" : "unset-topmost" });
  }

  async showWindow(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "show" });
  }

  async hideWindow(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "hide" });
  }

  async bringWindowToFront(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "bring-to-front" });
  }

  async getWindowInfo(workbookPath: string): Promise<unknown> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    return this.callTool("window", { session_id, action: "get-info" });
  }

  async setWindowState(workbookPath: string, state: "normal" | "minimized" | "maximized"): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "set-state", window_state: state });
  }

  async setWindowPosition(
    workbookPath: string,
    left: number,
    top: number,
    width: number,
    height: number
  ): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "set-position", left, top, width, height });
  }

  async arrangeWindow(workbookPath: string, preset: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "arrange", preset });
  }

  async setStatusBar(workbookPath: string, text: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "set-status-bar", text });
  }

  async clearStatusBar(workbookPath: string): Promise<void> {
    const session_id = await this.sessionStore.ensureSession(workbookPath);
    await this.callTool("window", { session_id, action: "clear-status-bar" });
  }

  // ============================================================
  // Macros
  // ============================================================

  async listMacros(workbookPath: string): Promise<Array<{ name: string; module: string; procedure: string }>> {
    const modules = await this.vbaList(workbookPath);
    const macros: Array<{ name: string; module: string; procedure: string }> = [];
    const re = /^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)\s*\(/gim;

    for (const mod of modules) {
      const code = await this.vbaView(workbookPath, mod.name);
      let m: RegExpExecArray | null;
      while ((m = re.exec(code)) !== null) {
        const procedure = m[1];
        macros.push({
          name: `${mod.name}.${procedure}`,
          module: mod.name,
          procedure,
        });
      }
      re.lastIndex = 0;
    }

    return macros;
  }

  // ============================================================
  // Checksum
  // ============================================================

  async getVbeCodeChecksum(workbookPath: string): Promise<string> {
    const modules = await this.vbaList(workbookPath);
    const chunks: string[] = [];
    for (const mod of modules) {
      const code = await this.vbaView(workbookPath, mod.name);
      chunks.push(`${mod.name}:${code.length}:${code}`);
    }
    return createHash("md5").update(chunks.join("\n")).digest("hex");
  }

  // ============================================================
  // Private helpers
  // ============================================================

  private toSnakeFormat(format: CellFormat): SnakeCellFormat {
    return {
      bold: format.bold,
      italic: format.italic,
      underline: format.underline,
      font_color: format.color,
      fill_color: format.backgroundColor,
      font_size: format.fontSize,
      number_format: format.numberFormat,
      horizontal_alignment: format.horizontalAlignment,
    };
  }

  private parseColor(color?: string): { red: number; green: number; blue: number } {
    if (!color) return { red: 0, green: 0, blue: 0 };
    const hex = color.replace("#", "");
    if (hex.length !== 6) return { red: 0, green: 0, blue: 0 };
    return {
      red: parseInt(hex.substring(0, 2), 16),
      green: parseInt(hex.substring(2, 4), 16),
      blue: parseInt(hex.substring(4, 6), 16),
    };
  }

  private shiftColumn(column: string, delta: number): string {
    let result = 0;
    for (const ch of column.toUpperCase()) {
      result = result * 26 + (ch.charCodeAt(0) - 64);
    }
    result += delta;
    let out = "";
    while (result > 0) {
      result--;
      out = String.fromCharCode((result % 26) + 65) + out;
      result = Math.floor(result / 26);
    }
    return out || "A";
  }
}
