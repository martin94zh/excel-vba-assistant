# vba-run

List and run VBA macros.

## Tools

- `excel_list_macros`
- `excel_run_macro`

## Parameters

### excel_list_macros

```json
{ "workbookId": "book" }
```

### excel_run_macro

```json
{
  "workbookId": "book",
  "macroName": "VBAProject.Module1.Main",
  "timeoutSeconds": 60,
  "captureResultRange": "A1"
}
```

- `macroName` (required): prefer fully qualified name `VBAProject.Module1.Procedure` or `Sheet1.Procedure`.
- `args` (optional): argument list for parameterized macros (currently limited support).
- `timeoutSeconds` (optional): default 45.
- `captureResultRange` (optional): read this range's `.Text` after execution and return it.

## Important

- Blocked when `VBE_READ_ONLY=true`.
- Only parameterless macros are fully supported in this version.
- If the macro shows a MsgBox/InputBox, the call blocks until timeout. Use `excel_list_dialogs` + `excel_click_dialog` / `excel_fill_dialog` to handle the popup before the timeout.
- Always confirm the macro exists via `excel_list_macros` first.
