# cell-read

Read cell values, formulas, and range metadata from a worksheet.

## Tool

`range`

## Values & Formulas

### Read values

```json
{
  "action": "get-values",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

- `rangeAddress` (required): any valid Excel range reference.
- `sheetName` (optional): defaults to the active worksheet.

### Read formulas

```json
{
  "action": "get-formulas",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

> Reading formulas works regardless of calculation mode.

## Range Discovery

### Get used range

```json
{
  "action": "get-used-range",
  "session_id": "abc123",
  "sheet_name": "Sheet1"
}
```

### Get current region

```json
{
  "action": "get-current-region",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
}
```

### Get range info

```json
{
  "action": "get-range-info",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

Returns address, dimensions, and other metadata.

## Search

### Find

```json
{
  "action": "find",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D100",
  "what": "Total",
  "matchCase": false,
  "matchEntireCell": false
}
```

### Replace

```json
{
  "action": "replace",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D100",
  "what": "OldText",
  "replacement": "NewText",
  "matchCase": false
}
```

## Sort

```json
{
  "action": "sort",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D100",
  "sortSpecs": [
    { "column": 1, "sortOrder": "ascending" },
    { "column": 3, "sortOrder": "descending" }
  ]
}
```

## Number Formats

### Get number formats

```json
{
  "action": "get-number-formats",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

## Hyperlinks

### List hyperlinks

```json
{
  "action": "list-hyperlinks",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

### Get hyperlink

```json
{
  "action": "get-hyperlink",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
}
```

## Notes

- Empty cells return `null` in the values matrix.
- Use a broad range first to discover data boundaries before creating charts or pivot tables.
