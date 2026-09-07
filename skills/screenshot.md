# screenshot

Capture ranges or worksheets as PNG images for visual verification.

## Tool

`screenshot`

## Parameters

### Capture range

```json
{
  "action": "capture",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:D20"
}
```

### Capture worksheet

```json
{
  "action": "capture-sheet",
  "session_id": "abc123",
  "sheet_name": "Dashboard"
}
```

## Returns

- MCP: image content as base64 PNG

## Notes

- Captures formatting, charts, and conditional formatting as rendered by Excel.
- Useful for verifying dashboard layout, chart appearance, and formatting results.
