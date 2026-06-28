# tab-format

Change worksheet tab color and visibility.

## Tools

- `excel_set_sheet_tab_format`

## Parameters

```json
{
  "workbookId": "book",
  "sheetName": "Sheet1",
  "color": "#FF0000",
  "visible": "Hidden"
}
```

## Options

- `color`: hex color for the sheet tab, e.g. `#FF0000`.
- `visible`: one of
  - `Visible` — normal visible tab
  - `Hidden` — hidden but can be unhidden via Excel UI
  - `VeryHidden` — deeply hidden, can only be restored by VBA/COM

## Important

- This tool is blocked when `VBE_READ_ONLY=true`.
- At least one of `color` or `visible` must be provided.
- A workbook must always have at least one visible sheet.
