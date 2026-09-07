# pivot-table

Create and manage PivotTables from ranges, Excel Tables, or the Data Model.

## Tool

`pivottable`

## Lifecycle

### List PivotTables

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Read PivotTable info

```json
{
  "action": "read",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

### Create PivotTable

```json
{
  "action": "create",
  "session_id": "abc123",
  "sourceType": "table",
  "source_name": "SalesData",
  "destinationSheet": "PivotSheet",
  "destinationRange": "A3",
  "pivotTableName": "SalesPivot"
}
```

- `sourceType`: `range` | `table` | `datamodel`
- `sourceName`: range address, table name, or Data Model table name
- `destinationSheet` / `destinationRange`: where to place the PivotTable
- `pivotTableName` (optional)

### Delete PivotTable

```json
{
  "action": "delete",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

## Field Management

### List fields

```json
{
  "action": "list-fields",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

### Add field

```json
{
  "action": "add-field",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Product",
  "area": "row",
  "position": 0
}
```

- `area`: `row` | `column` | `value` | `filter`
- `position` (optional): zero-based order in the area

### Remove field

```json
{
  "action": "remove-field",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Product",
  "area": "row"
}
```

## Field Configuration

### Set value field function

```json
{
  "action": "set-field-function",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Revenue",
  "functionName": "Sum"
}
```

- `functionName`: `Sum`, `Average`, `Count`, `CountNumbers`, `Max`, `Min`, `Product`, `StdDev`, `StdDevp`, `Var`, `Varp`

### Set field name

```json
{
  "action": "set-field-name",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Revenue",
  "new_name": "Total Revenue"
}
```

### Set field number format

```json
{
  "action": "set-field-number-format",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Revenue",
  "number_format": "$#,##0.00"
}
```

### Set field filter criteria

```json
{
  "action": "set-field-filter",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Region",
  "filterType": "values",
  "values": ["North", "South"]
}
```

### Sort field

```json
{
  "action": "sort-field",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Product",
  "sortOrder": "ascending"
}
```

## Calculated Fields (Regular PivotTables)

### List calculated fields

```json
{
  "action": "list-calculated-fields",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

### Create calculated field

```json
{
  "action": "create-calculated-field",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Commission",
  "formula": "=Revenue*0.1"
}
```

### Delete calculated field

```json
{
  "action": "delete-calculated-field",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Commission"
}
```

## Calculated Members (OLAP/Data Model PivotTables)

### List calculated members

```json
{
  "action": "list-calculated-members",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

### Create calculated member

```json
{
  "action": "create-calculated-member",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "memberName": "Premium Products",
  "mdx": "Aggregate({[Product].[Product].[Product A], [Product].[Product].[Product B]})"
}
```

### Delete calculated member

```json
{
  "action": "delete-calculated-member",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "memberName": "Premium Products"
}
```

## Layout & Formatting

### Set layout

```json
{
  "action": "set-layout",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "layout": "outline"
}
```

- `layout`: `table` | `outline`

### Set subtotals

```json
{
  "action": "set-subtotals",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "fieldName": "Product",
  "showSubtotals": true
}
```

### Set grand totals

```json
{
  "action": "set-grand-totals",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot",
  "showRowGrandTotals": true,
  "showColumnGrandTotals": true
}
```

## Data Operations

### Get PivotTable data

```json
{
  "action": "get-data",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

### Refresh PivotTable

```json
{
  "action": "refresh",
  "session_id": "abc123",
  "pivotTableName": "SalesPivot"
}
```

## Notes

- Data Model-based PivotTables support OLAP features; regular PivotTables support calculated fields.
- Refresh PivotTables after source data changes.
- Place PivotTables on a new sheet to avoid overwriting existing data.
