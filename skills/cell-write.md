# cell-write

Write, clear, copy, and format cell values.

## Tools

- `range`
- `range_format`

## Values & Formulas

### Set range values

```json
{
  "action": "set-values",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1",
  "values": [["Name", "Score"], ["Alice", 95], ["Bob", 88]]
}
```

### Set formulas

```json
{
  "action": "set-formulas",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "B2:B10",
  "formulas": [["=A2*10"], ["=A3*10"]]
}
```

## Clear

### Clear range contents

```json
{
  "action": "clear-contents",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

### Clear all

```json
{
  "action": "clear-all",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

### Clear formats

```json
{
  "action": "clear-formats",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

## Copy

### Copy range

```json
{
  "action": "copy",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "destinationSheet": "Sheet2",
  "destinationAddress": "A1"
}
```

### Copy values only

```json
{
  "action": "copy-values",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "destinationSheet": "Sheet2",
  "destinationAddress": "A1"
}
```

### Copy formulas only

```json
{
  "action": "copy-formulas",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "destinationSheet": "Sheet2",
  "destinationAddress": "A1"
}
```

## Structural Edits

### Insert/delete rows

```json
{
  "action": "insert-rows",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "row": 5,
  "count": 2
}
```

```json
{
  "action": "delete-rows",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "row": 5,
  "count": 2
}
```

### Insert/delete columns

```json
{
  "action": "insert-columns",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "column": 2,
  "count": 1
}
```

```json
{
  "action": "delete-columns",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "column": 2,
  "count": 1
}
```

### Insert/delete cells

```json
{
  "action": "insert-cells",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A5:A6",
  "shiftDirection": "down"
}
```

```json
{
  "action": "delete-cells",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A5:A6",
  "shiftDirection": "up"
}
```

## Number Formatting

### Set number format

```json
{
  "action": "set-number-format",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "number_format": "$#,##0.00"
}
```

### Set individual number formats

```json
{
  "action": "set-number-formats",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "numberFormats": [["yyyy-mm-dd", "0.00", "0.00%", "@"]]
}
```

## Visual Formatting

### Set cell format

```json
{
  "action": "set-format",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "format": {
    "bold": true,
    "color": "#FFFFFF",
    "backgroundColor": "#4472C4",
    "horizontal_alignment": "Center",
    "number_format": "0.00"
  }
}
```

### Format multiple ranges

```json
{
  "action": "format-ranges",
  "session_id": "abc123",
  "ranges": [
    {
      "sheet_name": "Sheet1",
      "range_address": "A1:D1",
      "format": { "bold": true, "backgroundColor": "#4472C4", "color": "#FFFFFF" }
    },
    {
      "sheet_name": "Sheet1",
      "range_address": "A2:D10",
      "format": { "number_format": "0.00" }
    }
  ]
}
```

### Built-in styles

```json
{
  "action": "set-style",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "styleName": "Heading 1"
}
```

```json
{
  "action": "get-style",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
}
```

## Merge & Protection

### Merge/unmerge cells

```json
{
  "action": "merge-cells",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D1"
}
```

```json
{
  "action": "unmerge-cells",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D1"
}
```

### Get merge info

```json
{
  "action": "get-merge-info",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D1"
}
```

### Cell lock status

```json
{
  "action": "set-cell-locked",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10",
  "locked": true
}
```

```json
{
  "action": "get-cell-locked",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
}
```

## Data Validation

### Add validation

```json
{
  "action": "add-validation",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:A10",
  "validation": {
    "type": "list",
    "formula": "North,South,East,West",
    "allowBlank": true,
    "showDropdown": true
  }
}
```

### Get validation

```json
{
  "action": "get-validation",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
}
```

### Remove validation

```json
{
  "action": "remove-validation",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:A10"
}
```

## Auto-Fit

```json
{
  "action": "auto-fit-columns",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

```json
{
  "action": "auto-fit-rows",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D10"
}
```

## Hyperlinks

### Add hyperlink

```json
{
  "action": "add-hyperlink",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1",
  "address": "https://example.com",
  "textToDisplay": "Open"
}
```

### Remove hyperlink

```json
{
  "action": "remove-hyperlink",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1"
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

- For multiple formatting changes, prefer one `range_format(set-format)` call with a large range over many small calls.
- For bulk writes (10+ cells), disable automatic recalculation first via `calculation_mode(set-mode, manual)`.
