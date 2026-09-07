# conditional-formatting

Add and clear conditional formatting rules.

## Tool

`conditionalformat`

## Parameters

### Add rule

```json
{
  "action": "add-rule",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:A100",
  "rule": {
    "type": "cellValue",
    "operator": "greaterThan",
    "formula": "50",
    "format": {
      "backgroundColor": "#FF0000",
      "color": "#FFFFFF"
    }
  }
}
```

Supported rule types:

- `cellValue`: cell value comparison (`>`, `<`, `=`, etc.)
- `expression`: formula-based rule
- `colorScale`: color scales
- `dataBar`: data bars
- `iconSet`: icon sets

### Clear rules

```json
{
  "action": "clear-rules",
  "session_id": "abc123",
  "sheet_name": "Sheet1",
  "range_address": "A1:A100"
}
```

## Notes

- Use conditional formatting for visual highlighting of trends, thresholds, and exceptions.
- For complex rules, build them incrementally and verify with `screenshot`.
