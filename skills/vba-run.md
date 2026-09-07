# vba-run

Run a VBA macro in the workbook.

## Tool

`vba`

## Parameters

```json
{
  "action": "run",
  "session_id": "abc123",
  "macroName": "Module1.Main"
}
```

- `macroName` (required): prefer fully qualified name `VBAProject.Module1.Procedure` or `Sheet1.Procedure`.

## Returns

- `success`: boolean
- `message`: execution result or error description

## Important

- Only parameterless macros are fully supported.
- Always confirm the macro exists via `vba(list)` + `vba(view)` first.
- After running a macro, check results via `range(get-values)` or the macro's own output.
