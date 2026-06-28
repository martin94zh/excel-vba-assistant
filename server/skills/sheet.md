# sheet

Manage worksheets in the workbook.

## Tools

- `excel_list_sheets`
- `excel_create_sheet`
- `excel_delete_sheet`
- `excel_rename_sheet`

## Parameters

### excel_list_sheets

```json
{ "workbookId": "book" }
```

### excel_create_sheet

```json
{
  "workbookId": "book",
  "sheetName": "NewSheet",
  "before": "Sheet1"
}
```

- `before` or `after` (optional): name of the reference sheet.

### excel_delete_sheet

```json
{
  "workbookId": "book",
  "sheetName": "OldSheet"
}
```

### excel_rename_sheet

```json
{
  "workbookId": "book",
  "oldName": "Sheet1",
  "newName": "Data"
}
```

## Important

- All write operations are blocked when `VBE_READ_ONLY=true`.
- Excel usually refuses to delete the last visible worksheet; create a replacement first if needed.
- Worksheet names in Excel are case-insensitive but displayed as entered.
