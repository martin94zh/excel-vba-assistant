# vba-write

Modify the VBAProject: update component code, create new components, or delete components.

## Tools

- `excel_update_vba_code`
- `excel_create_vba_component`
- `excel_delete_vba_component`

## Parameters

### excel_update_vba_code

```json
{
  "workbookId": "book",
  "componentName": "Module1",
  "code": "Sub Main()\n    MsgBox \"Hello\"\nEnd Sub"
}
```

- `componentName` (required): must already exist in VBE.
- `code` (required): clean VBA code. Do **not** include `VERSION`, `Attribute`, `Begin`, or `End` header lines.

### excel_create_vba_component

```json
{
  "workbookId": "book",
  "componentType": "standardModule",
  "componentName": "Module2",
  "code": "Sub NewProc()\nEnd Sub"
}
```

- `componentType` (required): `standardModule`, `classModule`, or `userForm`.
- `componentName` (required): English letter followed by letters/digits/underscores.
- `code` (optional): initial clean code.

### excel_delete_vba_component

```json
{
  "workbookId": "book",
  "componentName": "Module2"
}
```

After deleting a component with this tool, the local sync directory still contains the old file. Call `excel_sync_vbe_to_local` to remove it, or delete the local file and call `excel_sync_local_to_vbe` to delete the VBE component via sync.

## Important

- These tools are blocked when `VBE_READ_ONLY=true`.
- You cannot create or delete document modules (`Sheet1`, `ThisWorkbook`). Use `excel_update_vba_code` to change their code.
- Creating `userForm` via this tool only creates the code module. If you need the visual layout, create the form in VBE first, then update its code.
- Always read the existing code with `excel_get_vba_code` before overwriting, to avoid accidental loss.
