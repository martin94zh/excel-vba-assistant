# calculation

Control Excel calculation mode for performance and correctness.

## Tool

`calculation_mode`

## Parameters

### Get current calculation mode

```json
{
  "action": "get-mode",
  "session_id": "abc123"
}
```

### Set calculation mode

```json
{
  "action": "set-mode",
  "session_id": "abc123",
  "mode": "manual"
}
```

- `mode`: `automatic` | `manual` | `semi-automatic`

### Trigger recalculation

```json
{
  "action": "calculate",
  "session_id": "abc123",
  "scope": "workbook"
}
```

- `scope`: `workbook` | `sheet` | `range`
- `sheetName` (optional): required when scope is `sheet` or `range`
- `rangeAddress` (optional): required when scope is `range`

## Best Practice for Bulk Writes

When writing or updating many cells (10+), use this pattern to avoid recalculating after every cell:

```
1. calculation_mode(set-mode, manual)
2. Perform all writes (range set-values, set-formulas, etc.)
3. calculation_mode(calculate, workbook)
4. calculation_mode(set-mode, automatic)
```

> Reading formulas with `range(get-formulas)` works regardless of calculation mode.
