# workbook-inspect

Get a one-shot overview of the workbook: sheets, VBA resources, macros, and total code lines.

## Tool

`excel_inspect_workbook`

## Parameters

```json
{
  "workbookId": "book"
}
```

- `workbookId` (optional): omit to use the default workbook.

## When to Use

Call this first when exploring a new workbook, instead of making separate `excel_list_sheets`, `excel_list_resources`, and `excel_list_macros` calls.
