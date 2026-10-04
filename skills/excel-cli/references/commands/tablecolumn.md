### tablecolumn

Table column, filtering, and sorting operations for Excel Tables (ListObjects)

**Actions:** `apply-filter`, `clear-filters`, `get-filters`, `add-column`, `remove-column`, `rename-column`, `get-structured-reference`, `sort`, `sort-multi`, `get-column-number-format`, `set-column-number-format`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--table-name <TABLENAME>` | Name of the Excel table (required) |
| `--column-name <COLUMNNAME>` | Name of the column to filter (required for: apply-filter, add-column, remove-column, sort, get-column-number-format, set-column-number-format) (valid for: apply-filter, add-column, remove-column, get-structured-reference, sort, get-column-number-format, set-column-number-format) |
| `--options <OPTIONS>` | Native operator and its applicable criteria; nested keys remain camelCase (required for: apply-filter) (valid for: apply-filter) (JSON format) |
| `--position <POSITION>` | 1-based column position (optional, defaults to end of table) (valid for: add-column) |
| `--old-name <OLDNAME>` | Current column name (required for: rename-column) (valid for: rename-column) |
| `--new-name <NEWNAME>` | New column name (required for: rename-column) (valid for: rename-column) |
| `--region <REGION>` | Table region: 'Data', 'Headers', 'Totals', or 'All' Accepted values (case-insensitive): All, Data, Headers, Totals, ThisRow. (required for: get-structured-reference) (valid for: get-structured-reference) |
| `--ascending <ASCENDING>` | Sort order: true = ascending (A-Z, 0-9), false = descending (default: true) (valid for: sort) |
| `--sort-columns <SORTCOLUMNS>` | List of sort specifications: [{columnName: 'Col1', ascending: true}, ...] - applied in order (required for: sort-multi) (valid for: sort-multi) (JSON format) |
| `--format-code <FORMATCODE>` | Number format code in US locale (e.g., '#,##0.00', '0%', 'yyyy-mm-dd') (required for: set-column-number-format) (valid for: set-column-number-format) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
