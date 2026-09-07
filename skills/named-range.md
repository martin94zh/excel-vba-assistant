# named-range

Manage workbook parameters via named ranges.

## Tool

`namedrange`

## Lifecycle

### List named ranges

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Read named range

```json
{
  "action": "read",
  "session_id": "abc123",
  "name": "StartDate"
}
```

### Write named range

```json
{
  "action": "write",
  "session_id": "abc123",
  "name": "StartDate",
  "value": "2025-01-01"
}
```

### Create named range

```json
{
  "action": "create",
  "session_id": "abc123",
  "name": "ReportRange",
  "refersTo": "=Sheet1!$A$1:$D$100"
}
```

### Update named range

```json
{
  "action": "update",
  "session_id": "abc123",
  "name": "ReportRange",
  "refersTo": "=Sheet1!$A$1:$E$150"
}
```

### Delete named range

```json
{
  "action": "delete",
  "session_id": "abc123",
  "name": "ReportRange"
}
```

## Notes

- Hidden/internal Excel names are omitted from `list`.
- Large ranges return metadata without materializing values.
- Ideal for parameter-driven workbooks: update a named range → Power Query/connections refresh automatically.
