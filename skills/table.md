# table

Read and write Excel Tables (ListObjects).

## Tools

- `excel_list_tables`
- `excel_read_table`
- `excel_write_table`
- `excel_create_table`
- `excel_delete_table`

## Parameters

### excel_list_tables

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1"
}
```

Returns each table's `name`, `range`, `headerRowRange`, `dataBodyRange`, `rowCount`, `columnCount` and `style`.

### excel_read_table

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "tableName": "Table1",
  "includeHeaders": true
}
```

Returns `headers` and `rows`. Set `includeHeaders: false` to omit headers.

### excel_write_table

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "tableName": "Table1",
  "data": [["Alice", 95], ["Bob", 88]],
  "autoResize": true
}
```

Overwrites the table body (keeps headers). `autoResize` adjusts table size to match the data and adds/removes columns as needed.

### excel_create_table

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "tableName": "Table1",
  "address": "A1:D10",
  "hasHeaders": true,
  "styleName": "TableStyleMedium2"
}
```

Converts the range to a table. `address` must include the header row.

### excel_delete_table

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "tableName": "Table1",
  "clearDataOnly": false
}
```

Set `clearDataOnly: true` to clear table body data without removing the table structure.

## Important

- Write/delete/create tools are blocked when `VBE_READ_ONLY=true`.
- `excel_write_table` preserves the original header row; only the data body is replaced.
- When writing, ensure `data` rows have the same column count as the table header row.
