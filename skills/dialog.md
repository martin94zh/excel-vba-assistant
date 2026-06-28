# dialog

Inspect and interact with visible Excel/VBA dialogs.

## Tools

- `excel_list_dialogs`
- `excel_click_dialog`
- `excel_fill_dialog`

## Parameters

### excel_list_dialogs

```json
{}
```

Returns visible dialogs with `handle`, `title`, `text`, `buttons`, `hasInputField`, and a recommended `action`.

### excel_click_dialog

```json
{
  "handle": "0000000000012345",
  "action": "ok"
}
```

- `handle` (required): from `excel_list_dialogs` or `excel_run_macro` response `details.pendingDialog.handle`.
- `action` (optional): `auto`, `ok`, `确定`, `cancel`, `取消`, `yes`, `是`, `no`, `否`, `end`, `结束`, `close`, `关闭`, `ignore`, `忽略`, `retry`, `重试`, `continue`, `继续`, `debug`, `调试`.
- `buttonText` (optional): exact button text; takes precedence over `action`.

### excel_fill_dialog

```json
{
  "handle": "0000000000012345",
  "text": "2026-03-31",
  "submit": true
}
```

- `text` (required): text to enter into the dialog's input field.
- `submit` (optional): press Enter after filling.

## Interactive Macro Dialog Handling

When `excel_run_macro` is called with `interactive=true` (default), it pauses on ANY dialog and returns the dialog info directly in the response — no need to call `excel_list_dialogs` separately. The response includes `details.sessionId` and `details.pendingDialog`.

**Workflow:**
1. `excel_run_macro` → pauses, returns dialog info + `sessionId`
2. Analyze `details.pendingDialog` (text, buttons, hasInputField, kind)
3. Call `excel_click_dialog` or `excel_fill_dialog` with the handle from `pendingDialog`
4. Call `excel_resume_macro` with `sessionId` to continue
5. Repeat until macro completes

**Common dialog scenarios:**
- **Runtime error** (`kind: "vb_runtime_error"`): buttons usually [结束, 调试, 帮助]. Click "结束" to stop or "调试" to debug.
- **InputBox** (`hasInputField: true`): use `excel_fill_dialog` to enter text, then `excel_resume_macro`.
- **MsgBox** (`kind: "info"`): click "确定" or "取消" as appropriate.
- **Confirmation** (`kind: "confirmation"`): click "是"/"否" or "确定"/"取消".

## Important

- `excel_click_dialog` and `excel_fill_dialog` are blocked when `VBE_READ_ONLY=true`.
- The handle is a hexadecimal string; pass it exactly as returned.
- When handling dialogs during macro execution, always call `excel_resume_macro` after processing the dialog, otherwise the macro stays paused.
