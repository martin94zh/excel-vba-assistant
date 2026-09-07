# table

Create and manage Excel Tables (ListObjects). Excel Tables provide structured references, auto-expansion, and are required for Data Model/DAX workflows.

## Tool

`table`

## Lifecycle

### List tables

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Read table metadata

```json
{
  "action": "read",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

### Create table

```json
{
  "action": "create",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D100",
  "tableName": "SalesData",
  "styleName": "TableStyleMedium2"
}
```

- `rangeAddress` must include headers in the first row.
- `tableName` (optional): must start with a letter, no spaces.
- `styleName` (optional): built-in Excel table style name.

### Rename table

```json
{
  "action": "rename",
  "session_id": "abc123",
  "tableName": "SalesData",
  "new_name": "Sales2025"
}
```

### Resize table

```json
{
  "action": "resize",
  "session_id": "abc123",
  "tableName": "SalesData",
  "range_address": "A1:D150"
}
```

### Delete table

```json
{
  "action": "delete",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

## Styling & Formatting

### Apply table style

```json
{
  "action": "apply-style",
  "session_id": "abc123",
  "tableName": "SalesData",
  "styleName": "TableStyleMedium9"
}
```

### Toggle totals row

```json
{
  "action": "set-show-totals",
  "session_id": "abc123",
  "tableName": "SalesData",
  "showTotals": true
}
```

### Set column totals

```json
{
  "action": "set-column-totals",
  "session_id": "abc123",
  "tableName": "SalesData",
  "totals": {
    "Revenue": "Sum",
    "Quantity": "Average"
  }
}
```

### Get/set column number formats

```json
{
  "action": "get-column-number-formats",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

```json
{
  "action": "set-column-number-formats",
  "session_id": "abc123",
  "tableName": "SalesData",
  "formats": {
    "Revenue": "$#,##0.00",
    "Date": "yyyy-mm-dd"
  }
}
```

## Data Operations

### Read table data

```json
{
  "action": "get-data",
  "session_id": "abc123",
  "tableName": "SalesData",
  "visibleOnly": false
}
```

- `visibleOnly`: when `true`, returns only rows visible after filters.

### Append rows

```json
{
  "action": "append-rows",
  "session_id": "abc123",
  "tableName": "SalesData",
  "rows": [["2025-01-01", "Product A", 10], ["2025-01-02", "Product B", 20]]
}
```

### Add table to Data Model

```json
{
  "action": "add-to-datamodel",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

> Required before creating DAX measures that reference this table.

## DAX-Backed Tables

### Create table from DAX

```json
{
  "action": "create-from-dax",
  "session_id": "abc123",
  "sheet_name": "Sheet2",
  "range_address": "A1",
  "tableName": "DaxSummary",
  "dax": "EVALUATE SUMMARIZE(SalesData, SalesData[Product], \"Total\", SUM(SalesData[Revenue]))"
}
```

### Update DAX query

```json
{
  "action": "update-dax",
  "session_id": "abc123",
  "tableName": "DaxSummary",
  "dax": "EVALUATE SUMMARIZE(SalesData, SalesData[Region], \"Total\", SUM(SalesData[Revenue]))"
}
```

### Get DAX query

```json
{
  "action": "get-dax",
  "session_id": "abc123",
  "tableName": "DaxSummary"
}
```

## Filters

### Apply criteria filter

```json
{
  "action": "apply-filter",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Revenue",
  "operator": "greaterThan",
  "value": 1000
}
```

### Apply values filter

```json
{
  "action": "apply-filter-values",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Region",
  "values": ["North", "South"]
}
```

### Clear filters

```json
{
  "action": "clear-filters",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

### Get filter state

```json
{
  "action": "get-filter-state",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

## Column Management

### Add column

```json
{
  "action": "add-column",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Discount",
  "position": 3
}
```

### Remove column

```json
{
  "action": "remove-column",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Discount"
}
```

### Rename column

```json
{
  "action": "rename-column",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Revenue",
  "new_name": "SalesAmount"
}
```

## Structured References

### Get structured reference

```json
{
  "action": "get-structured-reference",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Revenue"
}
```

## Sorting

### Single-column sort

```json
{
  "action": "sort",
  "session_id": "abc123",
  "tableName": "SalesData",
  "columnName": "Revenue",
  "sortOrder": "descending"
}
```

### Multi-column sort

```json
{
  "action": "multi-column-sort",
  "session_id": "abc123",
  "tableName": "SalesData",
  "sortLevels": [
    { "columnName": "Region", "sortOrder": "ascending" },
    { "columnName": "Revenue", "sortOrder": "descending" }
  ]
}
```

## Notes

- Prefer creating a table over working with plain ranges for tabular data.
- Tables auto-expand when appending rows via `append-rows`.
- Data Model tables must exist before DAX measures can reference them.
