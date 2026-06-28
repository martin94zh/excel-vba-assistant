/**
 * Excel VBA Assistant - MCP 工具定义与分发
 *
 * 对齐任务文档第二十节，定义 21 个 excel_* 工具：
 *   excel_list_workbooks / excel_open_workbook
 *   excel_list_resources / excel_get_vba_code
 *   excel_update_vba_code / excel_create_vba_component / excel_delete_vba_component
 *   excel_list_macros / excel_run_macro
 *   excel_list_sheets / excel_read_range / excel_read_used_range
 *   excel_sync_vbe_to_local / excel_sync_local_to_vbe
 *
 * 危险操作（update / create / delete / run_macro / sync_local_to_vbe）
 * 在 ExcelVbaServiceClient 中通过 VBE_READ_ONLY 环境变量提供只读模式保护。
 */
import type { Tool } from "@modelcontextprotocol/sdk/types.js";

import {
  clearRangeSchema,
  clickDialogSchema,
  createSheetSchema,
  createTableSchema,
  createVbaComponentSchema,
  deleteSheetSchema,
  deleteTableSchema,
  deleteVbaComponentSchema,
  fillDialogSchema,
  getAllVbaCodeSchema,
  getVbaCodeSchema,
  inspectWorkbookSchema,
  listDialogsSchema,
  listMacrosSchema,
  listResourcesSchema,
  listSheetsSchema,
  listTablesSchema,
  listWorkbooksSchema,
  openWorkbookSchema,
  readRangeSchema,
  readTableSchema,
  readUsedRangeSchema,
  renameSheetSchema,
  runMacroSchema,
  setCellFormatSchema,
  setCellValueSchema,
  setRangeValuesSchema,
  setSheetTabFormatSchema,
  syncLocalToVbeSchema,
  syncVbeToLocalSchema,
  updateVbaCodeSchema,
  writeTableSchema,
} from "./schemas";
import { ExcelVbaServiceClient } from "./client/ExcelVbaServiceClient";

/** MCP 工具定义数组 */
export function buildToolDefinitions(): Tool[] {
  return [
    {
      name: "excel_list_workbooks",
      description: "列出当前 MCP Server 中已注册的所有 Excel 工作簿（包括启动时通过 VBE_FILE_PATH 注册的默认工作簿，以及通过 excel_open_workbook 添加的工作簿）。",
      inputSchema: listWorkbooksSchema,
    },
    {
      name: "excel_open_workbook",
      description: "注册一个新的 Excel 工作簿。文件必须已在 Excel 中打开（或可被 Excel 打开）。返回的 workbookId 将作为后续工具调用的参数。",
      inputSchema: openWorkbookSchema,
    },
    {
      name: "excel_inspect_workbook",
      description: "一次性获取工作簿全景信息：工作表列表（含 UsedRange）、所有 VBA 资源、所有宏（Sub/Function）、代码总行数。推荐在 AI 第一次接触工作簿时调用，避免多次 list_sheets / list_resources / list_macros。workbookId 可选。",
      inputSchema: inspectWorkbookSchema,
    },
    {
      name: "excel_list_resources",
      description: "列出指定工作簿中 VBAProject 的所有组件（模块、类模块、窗体、Excel 对象模块），含每个组件的名称、类型、代码行数。workbookId 可选。",
      inputSchema: listResourcesSchema,
    },
    {
      name: "excel_get_vba_code",
      description: "读取指定工作簿中某个 VBA 组件的完整代码。返回 componentName、componentType、code 等字段。workbookId 可选。",
      inputSchema: getVbaCodeSchema,
    },
    {
      name: "excel_get_all_vba_code",
      description: "一次性读取工作簿中所有 VBA 组件的完整代码。适合 AI 快速了解整个项目代码。workbookId 可选。",
      inputSchema: getAllVbaCodeSchema,
    },
    {
      name: "excel_update_vba_code",
      description: "【危险操作】将完整 VBA 代码写入指定组件（覆盖性写入）。若 VBE_READ_ONLY=true 则被拒绝。请确保 Excel 中已开启「信任对 VBA 项目对象模型的访问」。",
      inputSchema: updateVbaCodeSchema,
    },
    {
      name: "excel_create_vba_component",
      description: "【危险操作】在指定工作簿中创建新的 VBA 组件（标准模块/类模块/窗体）。若 VBE_READ_ONLY=true 则被拒绝。文档对象模块（Sheet/ThisWorkbook）不能通过此工具创建。",
      inputSchema: createVbaComponentSchema,
    },
    {
      name: "excel_delete_vba_component",
      description: "【危险操作】删除指定工作簿中的 VBA 组件。文档对象模块不能删除。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: deleteVbaComponentSchema,
    },
    {
      name: "excel_list_macros",
      description: "解析指定工作簿中所有组件的代码，列出 Sub/Function 过程，返回 module.procedure 形式的宏名列表。workbookId 可选。",
      inputSchema: listMacrosSchema,
    },
    {
      name: "excel_run_macro",
      description: "【危险操作】在已打开的 Excel 中运行一个 VBA 宏。建议传完整限定名（VBAProject.模块名.过程名 或 Sheet1.过程名）。运行期间会自动检测并点击常见弹窗（如运行时错误、确认框）。当前版本仅支持无参数宏。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: runMacroSchema,
    },
    {
      name: "excel_list_dialogs",
      description: "列出当前 Excel 进程中所有可见的 VBA / Excel 弹窗（MsgBox、InputBox、运行时错误等），返回每个弹窗的句柄、标题、文本、按钮列表和推荐操作。",
      inputSchema: listDialogsSchema,
    },
    {
      name: "excel_click_dialog",
      description: "【危险操作】根据 excel_list_dialogs 返回的句柄，模拟点击弹窗上的指定按钮。可传 action 关键词（ok/确定/cancel/取消/yes/是/no/否等）或精确的 buttonText。",
      inputSchema: clickDialogSchema,
    },
    {
      name: "excel_fill_dialog",
      description: "【危险操作】根据 excel_list_dialogs 返回的句柄，向弹窗中的输入框（InputBox / 文本框）写入文本。可设置 submit=true 在写入后按回车提交。",
      inputSchema: fillDialogSchema,
    },
    {
      name: "excel_list_sheets",
      description: "列出指定工作簿中的所有工作表，含名称、序号和 UsedRange 地址。workbookId 可选。",
      inputSchema: listSheetsSchema,
    },
    {
      name: "excel_read_range",
      description: "读取指定工作表的单元格区域数据（Value2）。sheetName 省略时使用第一个工作表，workbookId 可选。例如 address = \"A1:D10\"。",
      inputSchema: readRangeSchema,
    },
    {
      name: "excel_read_used_range",
      description: "读取指定工作表的 UsedRange 数据（含地址、行列数、值矩阵）。sheetName 省略时使用第一个工作表，workbookId 可选。",
      inputSchema: readUsedRangeSchema,
    },
    {
      name: "excel_sync_vbe_to_local",
      description: "将 VBE 中所有 VBA 组件导出到本地同步目录，并生成 workbook.json 清单。本地目录结构：模块/、类模块/、窗体/、Microsoft Excel 对象/（工作表 *.wks、工作簿 *.wbk）。",
      inputSchema: syncVbeToLocalSchema,
    },
    {
      name: "excel_sync_local_to_vbe",
      description: "【危险操作】将本地同步目录中的代码覆盖性写回 VBE：模块/*.bas、类模块/*.cls、窗体/*.frm（事件代码）、Microsoft Excel 对象/*.wks / *.wbk（事件代码）。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: syncLocalToVbeSchema,
    },
    {
      name: "excel_create_sheet",
      description: "【危险操作】在指定工作簿中新建工作表。可通过 before 或 after 指定插入位置。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: createSheetSchema,
    },
    {
      name: "excel_delete_sheet",
      description: "【危险操作】删除指定工作簿中的工作表。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: deleteSheetSchema,
    },
    {
      name: "excel_rename_sheet",
      description: "【危险操作】重命名指定工作簿中的工作表。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: renameSheetSchema,
    },
    {
      name: "excel_set_cell_value",
      description: "【危险操作】设置指定工作表单个单元格的值。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: setCellValueSchema,
    },
    {
      name: "excel_set_range_values",
      description: "【危险操作】从 startAddress 开始批量写入二维数组到单元格区域。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: setRangeValuesSchema,
    },
    {
      name: "excel_clear_range",
      description: "【危险操作】清空指定工作表的单元格区域内容。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: clearRangeSchema,
    },
    {
      name: "excel_set_cell_format",
      description: "【危险操作】设置指定单元格或区域的格式（字体、颜色、对齐、数字格式等）。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: setCellFormatSchema,
    },
    {
      name: "excel_list_tables",
      description: "列出指定工作表中的所有超级表（Excel Table / ListObject），含表名、区域、数据体区域、行列数、样式。sheetName 省略时使用第一个工作表，workbookId 可选。",
      inputSchema: listTablesSchema,
    },
    {
      name: "excel_read_table",
      description: "读取指定超级表的表头和数据行。sheetName 省略时使用第一个工作表，workbookId 可选。",
      inputSchema: readTableSchema,
    },
    {
      name: "excel_write_table",
      description: "【危险操作】向指定超级表写入数据（覆盖数据主体，保留表头）。支持自动调整表大小。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: writeTableSchema,
    },
    {
      name: "excel_create_table",
      description: "【危险操作】在指定工作表中将一个区域转换为超级表。address 需包含表头，如 A1:D10。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: createTableSchema,
    },
    {
      name: "excel_delete_table",
      description: "【危险操作】删除指定超级表；clearDataOnly=true 时仅清空数据。sheetName 省略时使用第一个工作表，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: deleteTableSchema,
    },
    {
      name: "excel_set_sheet_tab_format",
      description: "【危险操作】设置工作表页签格式：页签颜色（color）和可见性（Visible/Hidden/VeryHidden）。sheetName 必填，workbookId 可选。若 VBE_READ_ONLY=true 则被拒绝。",
      inputSchema: setSheetTabFormatSchema,
    },
  ];
}

interface CallContext {
  client: ExcelVbaServiceClient;
}

/** 单条工具调用的统一返回结构 */
export interface ToolCallResult {
  content: Array<{ type: "text"; text: string }>;
  isError?: boolean;
}

/** 工具调用分发器 */
export async function dispatchToolCall(
  name: string,
  args: Record<string, unknown> | undefined,
  ctx: CallContext
): Promise<ToolCallResult> {
  const a = args || {};
  const json = (data: unknown) => [{ type: "text" as const, text: JSON.stringify(data, null, 2) }];
  const text = (s: string) => [{ type: "text" as const, text: s }];
  const wbId = (v: unknown): string | undefined => (typeof v === "string" && v.trim() ? v.trim() : undefined);
  const sheetName = (v: unknown): string | undefined => (typeof v === "string" && v.trim() ? v.trim() : undefined);

  try {
    switch (name) {
      case "excel_list_workbooks": {
        const list = ctx.client.listWorkbooks();
        if (list.length === 0) {
          return { content: text("当前没有已注册的工作簿。请通过 excel_open_workbook 注册，或设置环境变量 VBE_FILE_PATH。") };
        }
        return { content: json({ workbooks: list, count: list.length }) };
      }
      case "excel_open_workbook": {
        const filePath = String(a.filePath || "").trim();
        if (!filePath) {
          return { content: text("错误: 缺少参数 filePath"), isError: true };
        }
        const workbookId = a.workbookId ? String(a.workbookId).trim() : undefined;
        const entry = ctx.client.openWorkbook(filePath, workbookId);
        return {
          content: json({
            success: true,
            workbookId: entry.workbookId,
            filePath: entry.filePath,
            localDir: entry.localDir,
            message: `工作簿已注册: ${entry.workbookId}`,
          }),
        };
      }
      case "excel_inspect_workbook": {
        const r = await ctx.client.inspectWorkbook(wbId(a.workbookId));
        return { content: json(r), isError: !r.success };
      }
      case "excel_list_resources": {
        const r = await ctx.client.listResources(wbId(a.workbookId));
        return { content: json(r), isError: !r.success };
      }
      case "excel_get_vba_code": {
        const r = await ctx.client.getVbaCode(wbId(a.workbookId), String(a.componentName));
        return { content: json(r), isError: !r.success };
      }
      case "excel_get_all_vba_code": {
        const r = await ctx.client.getAllVbaCode(wbId(a.workbookId));
        return { content: json(r), isError: !r.success };
      }
      case "excel_update_vba_code": {
        const r = await ctx.client.updateVbaCode(
          wbId(a.workbookId),
          String(a.componentName),
          typeof a.code === "string" ? a.code : ""
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_create_vba_component": {
        const r = await ctx.client.createVbaComponent(
          wbId(a.workbookId),
          String(a.componentType) as "standardModule" | "classModule" | "userForm",
          String(a.componentName),
          typeof a.code === "string" ? a.code : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_delete_vba_component": {
        const r = await ctx.client.deleteVbaComponent(wbId(a.workbookId), String(a.componentName));
        return { content: json(r), isError: !r.success };
      }
      case "excel_list_macros": {
        const r = await ctx.client.listMacros(wbId(a.workbookId));
        return { content: json(r), isError: !r.success };
      }
      case "excel_run_macro": {
        const r = await ctx.client.runMacro(
          wbId(a.workbookId),
          String(a.macroName),
          Array.isArray(a.args) ? a.args : undefined,
          typeof a.timeoutSeconds === "number" ? a.timeoutSeconds : undefined,
          typeof a.captureResultRange === "string" ? a.captureResultRange : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_list_dialogs": {
        const r = await ctx.client.listDialogs();
        return { content: json(r), isError: !r.success };
      }
      case "excel_click_dialog": {
        const r = await ctx.client.clickDialog(
          String(a.handle),
          typeof a.action === "string" ? a.action : undefined,
          typeof a.buttonText === "string" ? a.buttonText : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_fill_dialog": {
        const r = await ctx.client.fillDialog(
          String(a.handle),
          String(a.text),
          a.submit === true
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_list_sheets": {
        const r = await ctx.client.listSheets(wbId(a.workbookId));
        return { content: json(r), isError: !r.success };
      }
      case "excel_read_range": {
        const r = await ctx.client.readRange(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.address)
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_read_used_range": {
        const r = await ctx.client.readUsedRange(wbId(a.workbookId), sheetName(a.sheetName));
        return { content: json(r), isError: !r.success };
      }
      case "excel_sync_vbe_to_local": {
        const r = await ctx.client.syncVbeToLocal(
          wbId(a.workbookId),
          a.localDir ? String(a.localDir) : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_sync_local_to_vbe": {
        const r = await ctx.client.syncLocalToVbe(
          wbId(a.workbookId),
          a.localDir ? String(a.localDir) : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_create_sheet": {
        const r = await ctx.client.createSheet(
          wbId(a.workbookId),
          String(a.sheetName),
          a.before ? String(a.before) : undefined,
          a.after ? String(a.after) : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_delete_sheet": {
        const r = await ctx.client.deleteSheet(wbId(a.workbookId), String(a.sheetName));
        return { content: json(r), isError: !r.success };
      }
      case "excel_rename_sheet": {
        const r = await ctx.client.renameSheet(
          wbId(a.workbookId),
          String(a.oldName),
          String(a.newName)
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_set_cell_value": {
        const r = await ctx.client.setCellValue(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.address),
          a.value as unknown
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_set_range_values": {
        const r = await ctx.client.setRangeValues(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.startAddress),
          a.values as unknown[][]
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_clear_range": {
        const r = await ctx.client.clearRange(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.address)
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_set_cell_format": {
        const r = await ctx.client.setCellFormat(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.address),
          a.format as Record<string, unknown>
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_list_tables": {
        const r = await ctx.client.listTables(wbId(a.workbookId), sheetName(a.sheetName));
        return { content: json(r), isError: !r.success };
      }
      case "excel_read_table": {
        const r = await ctx.client.readTable(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.tableName),
          typeof a.includeHeaders === "boolean" ? a.includeHeaders : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_write_table": {
        const r = await ctx.client.writeTable(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.tableName),
          a.data as unknown[][],
          typeof a.autoResize === "boolean" ? a.autoResize : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_create_table": {
        const r = await ctx.client.createTable(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.tableName),
          String(a.address),
          typeof a.hasHeaders === "boolean" ? a.hasHeaders : undefined,
          a.styleName ? String(a.styleName) : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_delete_table": {
        const r = await ctx.client.deleteTable(
          wbId(a.workbookId),
          sheetName(a.sheetName),
          String(a.tableName),
          typeof a.clearDataOnly === "boolean" ? a.clearDataOnly : undefined
        );
        return { content: json(r), isError: !r.success };
      }
      case "excel_set_sheet_tab_format": {
        const r = await ctx.client.setSheetTabFormat(wbId(a.workbookId), String(a.sheetName), {
          color: a.color ? String(a.color) : undefined,
          visible: a.visible as "Visible" | "Hidden" | "VeryHidden" | undefined,
        });
        return { content: json(r), isError: !r.success };
      }
      default:
        return { content: text(`未知工具: ${name}`), isError: true };
    }
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    return { content: text(`错误: ${msg}`), isError: true };
  }
}
