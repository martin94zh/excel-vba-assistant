### sheet

Worksheet lifecycle management: create, rename, copy, delete, move, list sheets

**Actions:** `list`, `create`, `rename`, `copy`, `delete`, `move`, `copy-to-file`, `move-to-file`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--file-path <FILEPATH>` | Optional file path when batch contains multiple workbooks. If omitted, uses primary workbook. (valid for: list, create) |
| `--sheet <SHEETNAME>` | Exact name for the new worksheet. Names are not trimmed; blank or whitespace-only names are rejected. (required for: create, delete, move) (valid for: create, delete, move) |
| `--old-name <OLDNAME>` | Current name of the worksheet (required for: rename) (valid for: rename) |
| `--new-name <NEWNAME>` | Exact new worksheet name. Names are not trimmed; blank or whitespace-only names are rejected. (required for: rename) (valid for: rename) |
| `--source-name <SOURCENAME>` | Name of the source worksheet (required for: copy) (valid for: copy) |
| `--target-name <TARGETNAME>` | Exact name for the copied worksheet. Names are not trimmed; blank or whitespace-only names are rejected. (required for: copy) (valid for: copy) |
| `--before-sheet <BEFORESHEET>` | Optional: Name of sheet to position before (valid for: move, copy-to-file, move-to-file) |
| `--after-sheet <AFTERSHEET>` | Optional: Name of sheet to position after (valid for: move, copy-to-file, move-to-file) |
| `--source-file <SOURCEFILE>` | Full path to the source workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--source-sheet <SOURCESHEET>` | Name of the sheet to copy (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-file <TARGETFILE>` | Full path to the target workbook (required for: copy-to-file, move-to-file) (valid for: copy-to-file, move-to-file) |
| `--target-sheet-name <TARGETSHEETNAME>` | Optional: New name for the copied sheet (default: keeps original name) (valid for: copy-to-file) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
