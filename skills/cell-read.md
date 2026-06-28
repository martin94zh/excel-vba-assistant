# cell-read

Read cell values from a worksheet.

## Tools

- `excel_read_range`
- `excel_read_used_range`

## Parameters

### excel_read_range

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "address": "A1:D10"
}
```

- `address` (required): any valid Excel range reference.
- `sheetName` (optional): defaults to the first worksheet.

### excel_read_used_range

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1"
}
```

## Returns

- `address`: actual range address
- `rowCount`, `columnCount`
- `values`: 2D array of cell values

## Notes

- Use `excel_read_used_range` to discover data boundaries before creating charts or pivot tables.
- Empty cells return `null` in the values matrix.
