# sheet

Manage worksheets in the workbook.

## Tool

`worksheet`

## Lifecycle

### List sheets

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Create sheet

```json
{
  "action": "create",
  "session_id": "abc123",
  "name": "NewSheet",
  "before": "Sheet1"
}
```

- `before` or `after` (optional): name of the reference sheet.

### Rename sheet

```json
{
  "action": "rename",
  "session_id": "abc123",
  "old_name": "Sheet1",
  "new_name": "Data"
}
```

### Delete sheet

```json
{
  "action": "delete",
  "session_id": "abc123",
  "name": "OldSheet"
}
```

## Copy / Move

### Copy sheet

```json
{
  "action": "copy",
  "session_id": "abc123",
  "name": "Data",
  "before": "Sheet2"
}
```

### Move sheet

```json
{
  "action": "move",
  "session_id": "abc123",
  "name": "Data",
  "after": "Summary"
}
```

### Copy to another workbook

```json
{
  "action": "copy",
  "session_id": "abc123",
  "name": "Data",
  "target_file": "D:\\work\\Other.xlsx"
}
```

### Move to another workbook

```json
{
  "action": "move",
  "session_id": "abc123",
  "name": "Data",
  "target_file": "D:\\work\\Other.xlsx"
}
```

## Tab Colors

### Set tab color

```json
{
  "action": "set-tab-color",
  "session_id": "abc123",
  "name": "Data",
  "color": "#4472C4"
}
```

### Get tab color

```json
{
  "action": "get-tab-color",
  "session_id": "abc123",
  "name": "Data"
}
```

### Clear tab color

```json
{
  "action": "clear-tab-color",
  "session_id": "abc123",
  "name": "Data"
}
```

## Visibility

### Set visibility

```json
{
  "action": "set-visibility",
  "session_id": "abc123",
  "name": "Data",
  "visibility": "hidden"
}
```

- `visibility`: `visible` | `hidden` | `very-hidden`

### Get visibility

```json
{
  "action": "get-visibility",
  "session_id": "abc123",
  "name": "Data"
}
```

## Notes

- Excel usually refuses to delete the last visible worksheet; create a replacement first if needed.
- Worksheet names in Excel are case-insensitive but displayed as entered.
- `very-hidden` sheets are not visible in the Excel UI unhide dialog.
