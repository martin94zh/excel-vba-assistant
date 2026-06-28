# cell-write

Write and format cell values.

## Tools

- `excel_set_cell_value`
- `excel_set_range_values`
- `excel_clear_range`
- `excel_set_cell_format`

## Parameters

### excel_set_cell_value

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "address": "A1",
  "value": "Hello"
}
```

### excel_set_range_values

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "startAddress": "A1",
  "values": [["Name", "Score"], ["Alice", 95], ["Bob", 88]]
}
```

### excel_clear_range

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "address": "A1:D10"
}
```

### excel_set_cell_format

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "address": "A1:D10",
  "format": {
    "bold": true,
    "color": "#FFFFFF",
    "backgroundColor": "#4472C4",
    "horizontalAlignment": "Center",
    "numberFormat": "0.00"
  }
}
```

## Format Options

- `bold`, `italic`, `underline`
- `color`: font hex color
- `backgroundColor`: fill hex color
- `fontSize`: number
- `numberFormat`: e.g. `0.00`, `yyyy-mm-dd`
- `horizontalAlignment`: `Left`, `Center`, `Right`

## Important

- All tools in this document are blocked when `VBE_READ_ONLY=true`.
- For multiple formatting changes, prefer one `excel_set_cell_format` call with a large range over many small calls.
