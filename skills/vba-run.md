# vba-run

List and run VBA macros with interactive dialog handling.

## Tools

- `excel_list_macros`
- `excel_run_macro`
- `excel_resume_macro`

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
  "captureResultRange": "A1",
  "interactive": true
}
```

- `macroName` (required): prefer fully qualified name `VBAProject.Module1.Procedure` or `Sheet1.Procedure`.
- `args` (optional): argument list for parameterized macros (currently limited support).
- `timeoutSeconds` (optional): default 45.
- `captureResultRange` (optional): read this range's `.Text` after execution and return it.
- `autoFillInputs` (optional): text list to auto-fill InputBox dialogs in order.
- `interactive` (optional, default `true`): interactive dialog handling mode.

### excel_resume_macro

```json
{
  "sessionId": "abc123def456",
  "extendTimeoutSeconds": 30
}
```

- `sessionId` (required): from `excel_run_macro` response `details.sessionId`.
- `extendTimeoutSeconds` (optional): extend the macro timeout.

## Interactive Dialog Handling Workflow

When `interactive=true` (default), `excel_run_macro` pauses and returns immediately upon detecting any dialog (MsgBox, InputBox, runtime error, etc.). The response includes:

```json
{
  "success": false,
  "message": "宏执行期间检测到VBA 运行时错误/编译错误：...",
  "details": {
    "sessionId": "abc123",
    "kind": "vb_runtime_error",
    "pendingDialog": {
      "handle": "0000000000012345",
      "title": "Microsoft Visual Basic",
      "text": "运行时错误 '13': 类型不匹配",
      "buttons": ["结束", "调试", "帮助"],
      "hasInputField": false,
      "recommendedAction": "结束",
      "recommendedButtons": ["结束", "End", "确定", "OK"]
    }
  }
}
```

**AI workflow for handling dialogs:**

1. `excel_run_macro` → returns dialog info + `sessionId`
2. Analyze the dialog: read `pendingDialog.text` for error messages, `pendingDialog.buttons` for available buttons
3. Call `excel_click_dialog` or `excel_fill_dialog` to handle the dialog
   - For runtime errors: click "结束" (End) to stop, or "调试" (Debug) to debug
   - For InputBox: use `excel_fill_dialog` to enter text
   - For MsgBox: click "确定" (OK) or "取消" (Cancel) as appropriate
4. Call `excel_resume_macro` with `sessionId` to continue waiting for macro completion
5. Repeat if more dialogs appear, until macro completes or times out

## Important

- Blocked when `VBE_READ_ONLY=true`.
- Only parameterless macros are fully supported in this version.
- With `interactive=true` (default), ALL dialogs are returned for AI to handle — no automatic button clicking.
- With `interactive=false`, dialogs are auto-handled (runtime errors → End, errors → OK, etc.).
- Always confirm the macro exists via `excel_list_macros` first.
- Use `autoFillInputs` for InputBox dialogs when values are known in advance (bypasses interactive pause).
