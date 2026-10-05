### namedrange

Named ranges for formulas/parameters

**Actions:** `list`, `write`, `read`, `update`, `create`, `delete`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--name <NAME>` | Name of the named range (required for: write, read, update, create, delete) (valid for: write, read, update, create, delete) |
| `--value <VALUE>` | Value to set. Invariant numeric and Boolean strings become typed values; other input remains text. (required for: write) (valid for: write) |
| `--reference <REFERENCE>` | New cell reference (e.g., Sheet1!$A$1:$B$10) (required for: update, create) (valid for: update, create) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
