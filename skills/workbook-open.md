# workbook-open

Register an Excel workbook so that subsequent tools can target it.

## Tool

`excel_open_workbook`

## Parameters

```json
{
  "filePath": "D:\\work\\book.xlsm",
  "workbookId": "book"
}
```

- `filePath` (required): full path to the Excel file. The file must already be open in Excel, or Excel must be able to open it.
- `workbookId` (optional): ID used in later tool calls. Defaults to the file name without extension.

## Notes

- Supported extensions: `.xlsm`, `.xlsb`, `.xlam`, `.xls`, `.xlsx`.
- If the file is not already open, the server attempts to open it via Excel COM.
- A `localDir` can be associated with the workbook for sync operations.
