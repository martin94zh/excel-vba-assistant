# vba-sync

Export VBA code from Excel VBE to a local directory, or import local code back into VBE.

> These sync operations are an **internal plugin function**. They are performed by the plugin's Legacy COM mechanism, which combines `vba(list/view/import/update/delete)` calls internally to perform batch synchronization. **AI must NOT directly call the MCP `vba` tool to do batch import/export/sync.**

## What AI CAN do with the `vba` tool

AI may call the MCP `vba` tool only for **single-component or macro operations**, for example:

- `vba(action: 'run', procedure_name: 'MyMacro')` — execute a VBA procedure.
- `vba(action: 'view', module_name: 'Module1')` — read a single module's code.
- `vba(action: 'update', module_name: 'Module1', vba_code: '...')` — update a single module's code.

For **batch VBE ↔ local synchronization**, use the plugin commands below instead.

## Plugin Commands

- `Excel VBA: VBE → 本地 同步`
- `Excel VBA: 本地 → VBE 同步`

The VBE ↔ local synchronization is handled entirely by the plugin's Legacy COM mechanism. AI does not need to manage, trigger, or intervene in this process.

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

- `本地 → VBE` sync is blocked when the workbook is read-only.
- Exported files are UTF-8.
- Document modules (`.wks`, `.wbk`) only update existing code; they never create or delete sheets/workbooks.
- UserForm `.frm` files update existing forms' code only; visual layout is not created from local files.
- Before `本地 → VBE` sync, inspect `workbook.json` and the local files to ensure the target components exist in VBE.

## Deleting Components via Sync

These sync tools also perform **delete synchronization**:

- After `VBE → 本地`, local files whose components no longer exist in VBE are removed.
- After `本地 → VBE`, VBE components whose files no longer exist locally are removed, except for worksheet/workbook object modules (`Sheet*.wks`, `ThisWorkbook.wbk`).

To delete a component via sync:

1. Delete the local file (`模块/Old.bas`, `类模块/Old.cls`, or `窗体/Old.frm`).
2. Run the plugin's **本地 → VBE 同步** command.

> Danger: deleting files or components via sync is irreversible.
