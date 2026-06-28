# workbook-list

List workbooks currently registered in the MCP server.

## Tool

`excel_list_workbooks`

## Parameters

None.

## Returns

Array of registered workbooks with `workbookId`, `filePath`, and `localDir`.

If no workbook is registered, register one with `excel_open_workbook` or set `VBE_FILE_PATH` before starting the server.
