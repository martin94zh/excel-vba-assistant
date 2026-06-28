---
name: excel-vba-assistant
description: Use Excel VBA Assistant MCP tools to read, write, sync, and run VBA in Excel workbooks. Read the referenced operation document before calling an aggregated excel_* MCP tool.
---

# Excel VBA Assistant MCP

Use the `excel-vba-assistant` MCP server when the task involves **VBA code, VBAProject structure, macro execution, or VBE ↔ local sync**. For ordinary Excel workbook data operations, prefer the `excel-mcp` server instead.

## Workflow

1. Use `excel_list_workbooks` to confirm the target workbook is registered. If not, register it with `excel_open_workbook` or set `VBE_FILE_PATH` before starting the server.
2. Read the operation document for the task domain.
3. Call the matching aggregated MCP tool with the documented parameters.
4. After modifying VBA code, verify by calling `excel_get_vba_code` or `excel_list_resources`.
5. After running a macro, check results via `excel_read_range` or the macro's own output range.
6. Never call a write/delete/run/sync tool without reviewing the relevant operation document first.

## Tool Map

- `excel_list_workbooks` / `excel_open_workbook`: workbook registration. See `workbook-list.md`, `workbook-open.md`.
- `excel_inspect_workbook`: one-shot workbook overview (sheets, VBA resources, macros, code line counts). See `workbook-inspect.md`.
- `excel_list_resources`, `excel_get_vba_code`, `excel_get_all_vba_code`: read VBAProject structure and code. See `vba-read.md`.
- `excel_update_vba_code`, `excel_create_vba_component`, `excel_delete_vba_component`: modify VBAProject. See `vba-write.md`.
- `excel_list_macros`, `excel_run_macro`: discover and execute macros. See `vba-run.md`.
- `excel_sync_vbe_to_local`, `excel_sync_local_to_vbe`: export/import VBA between Excel VBE and local directory. See `vba-sync.md`.
- `excel_list_dialogs`, `excel_click_dialog`, `excel_fill_dialog`: inspect and interact with Excel/VBA dialogs. See `dialog.md`.
- `excel_list_sheets`, `excel_create_sheet`, `excel_delete_sheet`, `excel_rename_sheet`: worksheet management. See `sheet.md`.
- `excel_read_range`, `excel_read_used_range`: read cell values. See `cell-read.md`.
- `excel_set_cell_value`, `excel_set_range_values`, `excel_clear_range`, `excel_set_cell_format`: write cells and formatting. See `cell-write.md`.
- `excel_list_tables`, `excel_read_table`, `excel_write_table`, `excel_create_table`, `excel_delete_table`: Excel Table (ListObject) operations. See `table.md`.
- `excel_set_sheet_tab_format`: worksheet tab color and visibility. See `tab-format.md`.

## Important

- **workbookId is optional** for most tools. If omitted, the server uses the default workbook (from `VBE_FILE_PATH` or the most recent `excel_open_workbook`).
- **sheetName is optional** for cell/sheet tools. If omitted, the first worksheet is used.
- **Read-only mode**: set `VBE_READ_ONLY=true` to block all dangerous operations. Write/delete/run/sync tools return an error when read-only.
- **Dangerous operations** include: `excel_update_vba_code`, `excel_create_vba_component`, `excel_delete_vba_component`, `excel_run_macro`, `excel_click_dialog`, `excel_fill_dialog`, `excel_sync_local_to_vbe`, `excel_create_sheet`, `excel_delete_sheet`, `excel_rename_sheet`, `excel_create_table`, `excel_delete_table`, `excel_write_table`, `excel_set_sheet_tab_format`, all other `excel_set_*` and `excel_clear_range`.
- **Trust access**: Excel must have "Trust access to the VBA project object model" enabled, or all VBA read/write operations fail.
- **Delete synchronization**: `excel_sync_vbe_to_local` removes local files for components deleted in VBE; `excel_sync_local_to_vbe` removes VBE components whose local files are deleted. Worksheet/workbook object modules are never deleted by sync.
- **Document modules**: never try to create or delete worksheet/workbook object modules (`Sheet1`, `ThisWorkbook`). Only update their code via `excel_update_vba_code` or sync.
- **UserForms**: do not create UserForms via sync or `excel_create_vba_component` if you need to preserve the visual layout. Create them in VBE first, then update their code.
- **Sync directory structure**:
  ```
  localDir/
  ├─ workbook.json
  ├─ Microsoft Excel 对象/Sheet1.wks
  ├─ Microsoft Excel 对象/ThisWorkbook.wbk
  ├─ 模块/Module1.bas
  ├─ 类模块/Class1.cls
  └─ 窗体/UserForm1.frm
  ```
- **Macro names**: prefer fully qualified names such as `VBAProject.Module1.Main` or `Sheet1.Worksheet_Activate` to avoid name collisions.
- **Dialogs**: always call `excel_list_dialogs` first to get the exact `handle`, then use `excel_click_dialog` or `excel_fill_dialog`. The `handle` is a hexadecimal window handle string.
- **Macro timeout**: if a macro shows a MsgBox or InputBox, it will block until the timeout expires. Use dialog tools to dismiss or fill the popup before the timeout.
- **Encoding**: exported local files are UTF-8. VBE internally stores code as GBK; the server handles conversion automatically.
- **Batch reads**: use `excel_inspect_workbook` or `excel_get_all_vba_code` instead of multiple individual list/get calls.

## Examples

List registered workbooks:

```json
{ "name": "excel_list_workbooks" }
```

Register a workbook:

```json
{ "name": "excel_open_workbook", "arguments": { "filePath": "D:\\work\\book.xlsm" } }
```

Read the VBA read guide:

```json
{ "path": "vba-read.md" }
```
