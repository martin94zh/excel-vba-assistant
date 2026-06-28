# vba-sync

Export VBA code from Excel VBE to a local directory, or import local code back into VBE.

## Tools

- `excel_sync_vbe_to_local`
- `excel_sync_local_to_vbe`

## Parameters

### excel_sync_vbe_to_local

```json
{
  "workbookId": "book",
  "localDir": "D:\\work\\vba_sync"
}
```

- `localDir` (optional): defaults to `VBE_LOCAL_DIR` or the workbook's registered `localDir`.

### excel_sync_local_to_vbe

```json
{
  "workbookId": "book",
  "localDir": "D:\\work\\vba_sync"
}
```

## Local Directory Structure

```
localDir/
├─ workbook.json
├─ Microsoft Excel 对象/Sheet1.wks
├─ Microsoft Excel 对象/ThisWorkbook.wbk
├─ 模块/Module1.bas
├─ 类模块/Class1.cls
└─ 窗体/UserForm1.frm
```

## Important

- `excel_sync_local_to_vbe` is blocked when `VBE_READ_ONLY=true`.
- Exported files are UTF-8; the server converts from VBE's internal encoding automatically.
- Document modules (`.wks`, `.wbk`) only update existing code, they never create or delete sheets/workbooks.
- UserForm `.frm` files update existing forms' code only; visual layout is not created from local files.
- Before `excel_sync_local_to_vbe`, inspect `workbook.json` and the local files to ensure the target components exist in VBE.

## Deleting Components via Sync

These sync tools also perform **delete synchronization**:

- After `excel_sync_vbe_to_local`, local files whose components no longer exist in VBE are removed.
- After `excel_sync_local_to_vbe`, VBE components whose files no longer exist locally are removed, except for worksheet/workbook object modules (`Sheet*.wks`, `ThisWorkbook.wbk`).

To delete a component through MCP:

1. Delete the local file (`模块/Old.bas`, `类模块/Old.cls`, or `窗体/Old.frm`).
2. Call `excel_sync_local_to_vbe`.

Or use `excel_delete_vba_component` for single-component deletion, then call `excel_sync_vbe_to_local` to update the local directory.

> Danger: deleting files or components via sync is irreversible. Use `VBE_READ_ONLY=true` to prevent accidental changes.
