# vba-read

Read VBAProject structure and code without modifying anything.

## Tool

`vba`

## Parameters

### List all VBA components

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

Returns every VBComponent with `name` and `type` (`standardModule` / `classModule` / `userForm` / `worksheet` / `document`).

### Read a single component's code

```json
{
  "action": "view",
  "session_id": "abc123",
  "module_name": "Module1"
}
```

- `moduleName` (required): exact VBE component name.

### List available macros

There is no dedicated macro-listing tool. To discover macros, first call `vba(list)` to get all modules, then call `vba(view)` for each module and parse `Sub` / `Function` definitions.

## Notes

- Use `vba(list)` to get all components, then iterate with `vba(view)` instead of multiple individual calls when possible.
- Always verify a macro exists before running it.
