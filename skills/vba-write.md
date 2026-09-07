# vba-write

Modify the VBAProject: create new components, update existing component code, or delete components.

## Tool

`vba`

## Parameters

### Update existing component code

```json
{
  "action": "update",
  "session_id": "abc123",
  "module_name": "Module1",
  "code": "Sub Main()\n    MsgBox \"Hello\"\nEnd Sub"
}
```

- `moduleName` (required): must already exist in VBE.
- `code` (required): clean VBA code.

### Create a new component

```json
{
  "action": "import",
  "session_id": "abc123",
  "module_name": "Module2",
  "code": "Sub NewProc()\nEnd Sub",
  "type": "standardModule"
}
```

- `type` (required): `standardModule`, `classModule`, or `userForm`.
- `moduleName` (required): English letter followed by letters/digits/underscores.
- `code` (optional): initial code.

### Delete a component

```json
{
  "action": "delete",
  "session_id": "abc123",
  "module_name": "Module2"
}
```

After deleting a component, the local sync directory may still contain the old file. Call the plugin's sync command or delete the local file and sync back to update the directory.

## Important

- You cannot create or delete document modules (`Sheet1`, `ThisWorkbook`). Use `vba(update)` to change their code only.
- Creating `userForm` via `vba(import)` only creates the code module. If you need the visual layout, create the form in VBE first, then update its code.
- Always read the existing code with `vba(view)` before overwriting, to avoid accidental loss.
