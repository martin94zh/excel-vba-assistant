# dax

Work with the Excel Data Model and DAX measures.

## Tool

`datamodel`

## Tables

### List Data Model tables

```json
{
  "action": "list-tables",
  "session_id": "abc123"
}
```

### Read table info

```json
{
  "action": "read-table",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

### Rename Data Model table

```json
{
  "action": "rename-table",
  "session_id": "abc123",
  "tableName": "SalesData",
  "new_name": "Sales"
}
```

### Delete Data Model table

```json
{
  "action": "delete-table",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

## Columns

### List columns

```json
{
  "action": "list-columns",
  "session_id": "abc123",
  "tableName": "SalesData"
}
```

## Measures

### List measures

```json
{
  "action": "list-measures",
  "session_id": "abc123"
}
```

### Create measure

```json
{
  "action": "create-measure",
  "session_id": "abc123",
  "tableName": "SalesData",
  "measureName": "TotalRevenue",
  "dax": "SUM(SalesData[Revenue])",
  "format": "Currency"
}
```

- `format` (optional): `Currency`, `Percentage`, `Decimal`, `General`
- `formatDax` (optional): `true` to request remote DAX formatting.

### Update measure

```json
{
  "action": "update-measure",
  "session_id": "abc123",
  "measureName": "TotalRevenue",
  "dax": "SUM(SalesData[Revenue])",
  "format": "Currency"
}
```

### Delete measure

```json
{
  "action": "delete-measure",
  "session_id": "abc123",
  "measureName": "TotalRevenue"
}
```

## Model Info

### Read model info

```json
{
  "action": "read-info",
  "session_id": "abc123"
}
```

### List workbook connections

```json
{
  "action": "list-workbook-connections",
  "session_id": "abc123"
}
```

## Relationships

### List relationships

```json
{
  "action": "list-relationships",
  "session_id": "abc123"
}
```

### Read relationship

```json
{
  "action": "read-relationship",
  "session_id": "abc123",
  "relationshipId": "Rel1"
}
```

### Create relationship

```json
{
  "action": "create-relationship",
  "session_id": "abc123",
  "fromTable": "SalesData",
  "fromColumn": "ProductID",
  "toTable": "Products",
  "toColumn": "ProductID",
  "active": true
}
```

### Update relationship

```json
{
  "action": "update-relationship",
  "session_id": "abc123",
  "relationshipId": "Rel1",
  "active": false
}
```

### Delete relationship

```json
{
  "action": "delete-relationship",
  "session_id": "abc123",
  "relationshipId": "Rel1"
}
```

## Refresh & Query

### Refresh Data Model

```json
{
  "action": "refresh",
  "session_id": "abc123"
}
```

### Evaluate DAX query

```json
{
  "action": "evaluate",
  "session_id": "abc123",
  "dax": "EVALUATE SUMMARIZE(SalesData, SalesData[Product], \"Total\", SUM(SalesData[Revenue]))"
}
```

### Execute DMV metadata query

```json
{
  "action": "execute-dmv",
  "session_id": "abc123",
  "dmvQuery": "SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES"
}
```

## Prerequisites

- Tables must be added to the Data Model first via `table(add-to-datamodel)`.
- DAX calculated columns are not supported; use Excel UI for calculated columns.
- DAX formulas are preserved exactly by default, subject to Excel locale separator translation.
