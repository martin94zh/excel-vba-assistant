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

Returns visible dialogs with `handle`, `title`, `text`, `buttons`, and a recommended `action`.

### excel_click_dialog

```json
{
  "handle": "0000000000012345",
  "action": "ok"
}
```

- `handle` (required): from `excel_list_dialogs`.
- `action` (optional): `auto`, `ok`, `确定`, `cancel`, `取消`, `yes`, `是`, `no`, `否`, `end`, `结束`, `close`, `关闭`, `ignore`, `忽略`, `retry`, `重试`, `continue`, `继续`.
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

## Important

- `excel_click_dialog` and `excel_fill_dialog` are blocked when `VBE_READ_ONLY=true`.
- Always call `excel_list_dialogs` first to obtain the current `handle`.
- The handle is a hexadecimal string; pass it exactly as returned.
