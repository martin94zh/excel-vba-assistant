# connection

Manage external data connections.

## Tool

`connection`

## Lifecycle

### List connections

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### View connection

```json
{
  "action": "view",
  "session_id": "abc123",
  "connectionName": "SalesDB"
}
```

### Create connection

```json
{
  "action": "create",
  "session_id": "abc123",
  "connectionName": "SalesDB",
  "connectionType": "OLEDB",
  "connectionString": "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\\Data\\Sales.accdb;",
  "commandText": "SELECT * FROM Sales"
}
```

- `connectionType`: `OLEDB`, `ODBC`
- Requires installed provider/driver (e.g. `Microsoft.ACE.OLEDB.16.0`).
- TEXT/WEB connections automatically redirect to Power Query for reliable imports.

### Test connection

```json
{
  "action": "test",
  "session_id": "abc123",
  "connectionName": "SalesDB"
}
```

### Refresh connection

```json
{
  "action": "refresh",
  "session_id": "abc123",
  "connectionName": "SalesDB"
}
```

### Delete connection

```json
{
  "action": "delete",
  "session_id": "abc123",
  "connectionName": "SalesDB"
}
```

## Load & Properties

### Load to worksheet

```json
{
  "action": "load-to",
  "session_id": "abc123",
  "connectionName": "SalesDB",
  "sheet_name": "Data",
  "range_address": "A1"
}
```

### Get properties

```json
{
  "action": "get-properties",
  "session_id": "abc123",
  "connectionName": "SalesDB"
}
```

### Set properties

```json
{
  "action": "set-properties",
  "session_id": "abc123",
  "connectionName": "SalesDB",
  "connectionString": "...",
  "commandText": "SELECT * FROM Sales WHERE Year = 2025"
}
```

## Notes

- Refreshing a connection may trigger external network/database queries.
- Power Query connections are handled atomically by `powerquery`.
