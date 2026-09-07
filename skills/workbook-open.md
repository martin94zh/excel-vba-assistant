# workbook-open

Open, create, test, and close Excel workbooks.

## Tool

`file`

## Open a workbook

```json
{
  "action": "open",
  "path": "D:\\work\\book.xlsm"
}
```

- `path` (required): full path to the Excel file.

### Returns

- `session_id`: a string identifier required by all subsequent tool calls.

## Create empty workbook

```json
{
  "action": "create",
  "path": "D:\\work\\new.xlsx",
  "format": "xlsx"
}
```

- `format`: `xlsx` | `xlsm`

## Test workbook access

```json
{
  "action": "test",
  "path": "D:\\work\\book.xlsx"
}
```

- Returns `isIrmProtected` flag for IRM/AIP-protected files.

## List active sessions

```json
{
  "action": "list"
}
```

## Close session

```json
{
  "action": "close",
  "session_id": "abc123",
  "save": true
}
```

- `save` (optional): `true` to save changes before closing.

## Close workbook without quitting Excel

```json
{
  "action": "close-workbook",
  "session_id": "abc123",
  "save": true
}
```

## Notes

- Supported extensions: `.xlsm`, `.xlsb`, `.xlam`, `.xls`, `.xlsx`.
- **If the file is already open in Excel**, `open` attaches to the running instance and returns a session — it does NOT start a second Excel process.
- If the file is not already open, the server attempts to open it via Excel COM.
- IRM/AIP-protected files are opened read-only with Excel visible for credential authentication.
- Always pass the returned `session_id` to all subsequent tools.
- Call `file(close)` when done to persist changes and release the session.
- **`file(list)` only shows sessions opened through this MCP server.** It does NOT detect manually opened Excel files. Always use `file(open, path)` to connect.
