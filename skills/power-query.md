# power-query

Manage Power Query queries and M code.

## Tool

`powerquery`

## Discovery

### List queries

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### View M code

```json
{
  "action": "view",
  "session_id": "abc123",
  "queryName": "SalesQuery"
}
```

## Evaluate First (Best Practice)

Always test M code before creating a persistent query.

```json
{
  "action": "evaluate",
  "session_id": "abc123",
  "mCode": "let Source = Excel.CurrentWorkbook(){[Name=\"SalesData\"]}[Content] in Source"
}
```

> Evaluating first catches syntax errors and missing sources without polluting the workbook with broken queries.

## Lifecycle

### Create query

```json
{
  "action": "create",
  "session_id": "abc123",
  "queryName": "SalesQuery",
  "mCode": "let Source = Excel.CurrentWorkbook(){[Name=\"SalesData\"]}[Content] in Source",
  "loadDestination": "Sheet2"
}
```

- `loadDestination` (optional): target sheet name; omit to load to Data Model only.
- `formatMCode` (optional): `true` to request remote formatting.

### Update query

```json
{
  "action": "update",
  "session_id": "abc123",
  "queryName": "SalesQuery",
  "mCode": "...",
  "autoRefresh": true
}
```

### Rename query

```json
{
  "action": "rename",
  "session_id": "abc123",
  "queryName": "SalesQuery",
  "new_name": "SalesQueryV2"
}
```

### Delete query

```json
{
  "action": "delete",
  "session_id": "abc123",
  "queryName": "SalesQuery"
}
```

## Refresh

### Refresh one query

```json
{
  "action": "refresh",
  "session_id": "abc123",
  "queryName": "SalesQuery"
}
```

### Refresh all queries

```json
{
  "action": "refresh-all",
  "session_id": "abc123"
}
```

## Load Configuration

### Load to destination

```json
{
  "action": "load-to",
  "session_id": "abc123",
  "queryName": "SalesQuery",
  "destination": "Sheet2",
  "refresh": true
}
```

### Get load config

```json
{
  "action": "get-load-config",
  "session_id": "abc123",
  "queryName": "SalesQuery"
}
```

### Unload query data

```json
{
  "action": "unload",
  "session_id": "abc123",
  "queryName": "SalesQuery"
}
```

> Unload removes loaded data but keeps the query definition.

## Notes

- M code is preserved exactly by default.
- `formatMCode=true` sends M code to a remote formatter and adds latency; if formatting fails, the original M code is saved unchanged.
- Power Query operations may open credential dialogs for external data sources.
- TEXT/WEB connections automatically redirect to Power Query for reliable imports.
