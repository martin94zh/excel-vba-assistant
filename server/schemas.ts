/**
 * Excel VBA Assistant - MCP 工具输入 Schema 定义
 *
 * 对齐任务文档第二十节，定义所有 MCP 工具的输入参数 schema。
 * 危险操作（修改 VBA、删除组件、运行宏、本地 → VBE 同步）通过 readOnly 全局开关
 * 或工具自身在执行前检查 VBE_READ_ONLY 环境变量来提供只读模式。
 */

const OBJ = "object" as const;
const STR = "string";
const NUM = "number";
const BOOL = "boolean";
const ARR = "array" as const;

const WORKBOOK_ID_DESC = "可选。工作簿 ID；省略时使用默认工作簿（由 VBE_FILE_PATH 或最近一次 excel_open_workbook 指定）。";
const SHEET_NAME_DESC = "可选。工作表名；省略时使用第一个工作表。";

/** excel_list_workbooks */
export const listWorkbooksSchema = {
  type: OBJ,
  properties: {},
};

/** excel_open_workbook */
export const openWorkbookSchema = {
  type: OBJ,
  properties: {
    filePath: {
      type: STR,
      description: "Excel 文件完整路径（支持 xlsm/xlsb/xlam/xlsx/xls）",
    },
    workbookId: {
      type: STR,
      description: "可选。后续工具调用引用此工作簿时使用的 ID；不填则使用文件名（不含扩展名）",
    },
  },
  required: ["filePath"],
};

/** excel_inspect_workbook（全景信息） */
export const inspectWorkbookSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
  },
  required: [],
};

/** excel_list_resources */
export const listResourcesSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
  },
  required: [],
};

/** excel_get_vba_code */
export const getVbaCodeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    componentName: { type: STR, description: "VBA 组件名，例如 Module1 / Sheet1 / ThisWorkbook / UserForm1" },
  },
  required: ["componentName"],
};

/** excel_get_all_vba_code */
export const getAllVbaCodeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
  },
  required: [],
};

/** excel_update_vba_code（危险：修改 VBA） */
export const updateVbaCodeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    componentName: { type: STR, description: "目标组件名（必须已存在于 VBE 中）" },
    code: { type: STR, description: "要写入的完整 VBA 代码（不要包含 VERSION / Attribute 行）" },
  },
  required: ["componentName", "code"],
};

/** excel_create_vba_component（危险：新增组件） */
export const createVbaComponentSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    componentType: {
      type: STR,
      description: "组件类型：standardModule(1) / classModule(2) / userForm(3)",
      enum: ["standardModule", "classModule", "userForm"],
    },
    componentName: { type: STR, description: "新组件名（英文开头，仅含字母数字下划线）" },
    code: { type: STR, description: "可选。新组件的初始代码（不要包含 VERSION / Attribute 行）" },
  },
  required: ["componentType", "componentName"],
};

/** excel_delete_vba_component（危险：删除组件） */
export const deleteVbaComponentSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    componentName: { type: STR, description: "要删除的组件名" },
  },
  required: ["componentName"],
};

/** excel_list_macros */
export const listMacrosSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
  },
  required: [],
};

/** excel_run_macro（危险：执行宏） */
export const runMacroSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    macroName: {
      type: STR,
      description: "宏名。建议传完整限定名：VBAProject.模块名.过程名 或 Sheet1.过程名 / ThisWorkbook.过程名",
    },
    args: {
      type: ARR,
      description: "可选。宏参数列表（仅支持基础类型）",
      items: {},
    },
    timeoutSeconds: {
      type: NUM,
      description: "可选。宏执行超时秒数，默认 45 秒。如果宏弹出 MsgBox / InputBox 等模态对话框，会触发超时。",
    },
    captureResultRange: {
      type: STR,
      description: "可选。宏执行后读取哪个单元格/区域的 Text 作为结果返回，例如 A1 或 Sheet1!B2。",
    },
  },
  required: ["macroName"],
};

/** excel_list_dialogs */
export const listDialogsSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: "保留字段，当前版本可忽略" },
  },
  required: [],
};

/** excel_click_dialog（危险：模拟点击弹窗） */
export const clickDialogSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: "保留字段，当前版本可忽略" },
    handle: { type: STR, description: "弹窗句柄，由 excel_list_dialogs 返回" },
    action: {
      type: STR,
      description: "动作关键词：auto(自动)/ok/确定/cancel/取消/yes/是/no/否/end/结束/close/关闭/ignore/忽略/retry/重试/continue/继续",
    },
    buttonText: { type: STR, description: "精确按钮文本，例如 确定。优先于 action。" },
  },
  required: ["handle"],
};

/** excel_fill_dialog（危险：向弹窗输入框写入文本） */
export const fillDialogSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: "保留字段，当前版本可忽略" },
    handle: { type: STR, description: "弹窗句柄，由 excel_list_dialogs 返回" },
    text: { type: STR, description: "要写入输入框的文本，例如 2026-03-31" },
    submit: { type: BOOL, description: "可选。写入后是否按回车提交，默认 false。" },
  },
  required: ["handle", "text"],
};

/** excel_list_sheets */
export const listSheetsSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
  },
  required: [],
};

/** excel_read_range */
export const readRangeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    address: { type: STR, description: "单元格区域，例如 A1:D10 / A:A / 1:5" },
  },
  required: ["address"],
};

/** excel_read_used_range */
export const readUsedRangeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
  },
  required: [],
};

/** excel_sync_vbe_to_local */
export const syncVbeToLocalSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    localDir: { type: STR, description: "可选。本地同步目录；不填则使用环境变量 VBE_LOCAL_DIR" },
  },
  required: [],
};

/** excel_sync_local_to_vbe（危险：覆盖性同步） */
export const syncLocalToVbeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    localDir: { type: STR, description: "可选。本地同步目录；不填则使用环境变量 VBE_LOCAL_DIR" },
  },
  required: [],
};

/** excel_create_sheet（危险：新增工作表） */
export const createSheetSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: "新工作表名称" },
    before: { type: STR, description: "可选。插入到指定工作表之前" },
    after: { type: STR, description: "可选。插入到指定工作表之后" },
  },
  required: ["sheetName"],
};

/** excel_delete_sheet（危险：删除工作表） */
export const deleteSheetSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR },
  },
  required: ["sheetName"],
};

/** excel_rename_sheet（危险：重命名工作表） */
export const renameSheetSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    oldName: { type: STR },
    newName: { type: STR },
  },
  required: ["oldName", "newName"],
};

/** excel_set_cell_value（危险：写入单元格） */
export const setCellValueSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    address: { type: STR, description: "单元格地址，如 A1" },
    value: { description: "单元格值（字符串、数字、布尔或 null）" },
  },
  required: ["address", "value"],
};

/** excel_set_range_values（危险：批量写入区域） */
export const setRangeValuesSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    startAddress: { type: STR, description: "起始单元格，如 A1" },
    values: {
      type: ARR,
      description: "二维数组，例如 [[\"A\",\"B\"],[1,2]]",
      items: { type: ARR, items: {} },
    },
  },
  required: ["startAddress", "values"],
};

/** excel_clear_range（危险：清空区域） */
export const clearRangeSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    address: { type: STR, description: "要清空的区域，如 A1:D10" },
  },
  required: ["address"],
};

/** excel_set_cell_format（危险：设置格式） */
export const setCellFormatSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    address: { type: STR, description: "单元格或区域，如 A1 或 A1:D10" },
    format: {
      type: OBJ,
      description: "格式对象",
      properties: {
        bold: { type: BOOL },
        italic: { type: BOOL },
        underline: { type: BOOL },
        color: { type: STR, description: "字体颜色，如 #FF0000" },
        backgroundColor: { type: STR, description: "背景颜色，如 #FFFF00" },
        fontSize: { type: NUM },
        numberFormat: { type: STR, description: "数字格式，如 0.00 / yyyy-mm-dd" },
        horizontalAlignment: { type: STR, enum: ["Left", "Center", "Right"] },
      },
    },
  },
  required: ["address", "format"],
};

/** excel_list_tables */
export const listTablesSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
  },
};

/** excel_read_table */
export const readTableSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    tableName: { type: STR, description: "超级表名称" },
    includeHeaders: { type: BOOL, description: "是否包含表头，默认 true" },
  },
  required: ["tableName"],
};

/** excel_write_table */
export const writeTableSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    tableName: { type: STR, description: "超级表名称" },
    data: {
      type: ARR,
      description: "二维数组数据，不包含表头",
      items: { type: ARR, items: {} },
    },
    autoResize: { type: BOOL, description: "是否自动调整表大小以匹配数据，默认 true" },
  },
  required: ["tableName", "data"],
};

/** excel_create_table */
export const createTableSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    tableName: { type: STR, description: "超级表名称" },
    address: { type: STR, description: "表区域，如 A1:D10（包含表头）" },
    hasHeaders: { type: BOOL, description: "首行是否为表头，默认 true" },
    styleName: { type: STR, description: "可选。表样式名称，如 TableStyleMedium2" },
  },
  required: ["tableName", "address"],
};

/** excel_delete_table */
export const deleteTableSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: SHEET_NAME_DESC },
    tableName: { type: STR, description: "超级表名称" },
    clearDataOnly: { type: BOOL, description: "仅清空数据而不删除表结构，默认 false" },
  },
  required: ["tableName"],
};

/** excel_set_sheet_tab_format */
export const setSheetTabFormatSchema = {
  type: OBJ,
  properties: {
    workbookId: { type: STR, description: WORKBOOK_ID_DESC },
    sheetName: { type: STR, description: "目标工作表名" },
    color: { type: STR, description: "页签颜色，如 #FF0000" },
    visible: { type: STR, enum: ["Visible", "Hidden", "VeryHidden"], description: "可见性：Visible 可见 / Hidden 隐藏（可通过取消隐藏恢复）/ VeryHidden 深度隐藏（需 VBA 恢复）" },
  },
  required: ["sheetName"],
};

export type ToolSchemaMap = {
  excel_list_workbooks: typeof listWorkbooksSchema;
  excel_open_workbook: typeof openWorkbookSchema;
  excel_inspect_workbook: typeof inspectWorkbookSchema;
  excel_list_resources: typeof listResourcesSchema;
  excel_get_vba_code: typeof getVbaCodeSchema;
  excel_get_all_vba_code: typeof getAllVbaCodeSchema;
  excel_update_vba_code: typeof updateVbaCodeSchema;
  excel_create_vba_component: typeof createVbaComponentSchema;
  excel_delete_vba_component: typeof deleteVbaComponentSchema;
  excel_list_macros: typeof listMacrosSchema;
  excel_run_macro: typeof runMacroSchema;
  excel_list_dialogs: typeof listDialogsSchema;
  excel_click_dialog: typeof clickDialogSchema;
  excel_fill_dialog: typeof fillDialogSchema;
  excel_list_sheets: typeof listSheetsSchema;
  excel_read_range: typeof readRangeSchema;
  excel_read_used_range: typeof readUsedRangeSchema;
  excel_sync_vbe_to_local: typeof syncVbeToLocalSchema;
  excel_sync_local_to_vbe: typeof syncLocalToVbeSchema;
  excel_create_sheet: typeof createSheetSchema;
  excel_delete_sheet: typeof deleteSheetSchema;
  excel_rename_sheet: typeof renameSheetSchema;
  excel_set_cell_value: typeof setCellValueSchema;
  excel_set_range_values: typeof setRangeValuesSchema;
  excel_clear_range: typeof clearRangeSchema;
  excel_set_cell_format: typeof setCellFormatSchema;
  excel_list_tables: typeof listTablesSchema;
  excel_read_table: typeof readTableSchema;
  excel_write_table: typeof writeTableSchema;
  excel_create_table: typeof createTableSchema;
  excel_delete_table: typeof deleteTableSchema;
  excel_set_sheet_tab_format: typeof setSheetTabFormatSchema;
};
