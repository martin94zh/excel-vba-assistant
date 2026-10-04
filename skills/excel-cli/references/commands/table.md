### table

Excel Tables (ListObjects) - lifecycle and data operations

**Actions:** `list`, `preflight`, `create`, `rename`, `delete`, `read`, `resize`, `toggle-totals`, `set-column-total`, `append`, `get-data`, `set-style`, `add-to-data-model`, `create-from-dax`, `update-dax`, `get-dax`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Name of the worksheet containing the proposed table (required for: preflight, create, create-from-da x) (valid for: preflight, create, create-from-da x) |
| `--table-name <TABLENAME>` | Name for the proposed table (must be unique in workbook) (required for: preflight, create, rename, delete, read, resize, toggle-totals, set-column-tot al, append, get-data, set-style, add-to-data-mo del, create-from-da x, update-dax, get-dax) (valid for: preflight, create, rename, delete, read, resize, toggle-totals, set-column-tot al, append, get-data, set-style, add-to-data-mo del, create-from-da x, update-dax, get-dax) |
| `--range <RANGEADDRESS>` | Cell range address, or one cell to expand to its CurrentRegion (required for: preflight, create) (valid for: preflight, create) |
| `--has-headers <HASHEADERS>` | True if the first row contains column headers (default: true) (valid for: preflight, create) |
| `--table-style <TABLESTYLE>` | Table style name (e.g., 'TableStyleMed ium2', 'TableStyleLig ht1'). Optional. (required for: set-style) (valid for: create, set-style) |
| `--new-name <NEWNAME>` | New name for the table (must be unique in workbook) (required for: rename) (valid for: rename) |
| `--new-range <NEWRANGE>` | New range address (e.g., 'A1:F20') (required for: resize) (valid for: resize) |
| `--show-totals <SHOWTOTALS>` | True to show totals row, false to hide (required for: toggle-totals) (valid for: toggle-totals) |
| `--column-name <COLUMNNAME>` | Name of the column to set total function on (required for: set-column-tot al) (valid for: set-column-tot al) |
| `--total-function <TOTALFUNCTION>` | Totals function name: Sum, Count, Average, Min, Max, CountNums, StdDev, Var, None (required for: set-column-tot al) (valid for: set-column-tot al) |
| `--rows <ROWS>` | 2D array of row data to append - every row must match the table's column count and column order. Optional if rowsFile is provided. (valid for: append) (JSON format) |
| `--rows-file <ROWSFILE>` | Path to a JSON or CSV file containing the rows to append. Every row must match the table's column count. JSON: 2D array. CSV: rows/columns. Alternative to inline rows parameter. (valid for: append) |
| `--visible-only <VISIBLEONLY>` | True to return only visible (non-filtered) rows; false for all rows (default: false) (valid for: get-data) |
| `--strip-bracket-column-names <STRIPBRACKETCOLUMNNAMES>` | When true, renames source table columns that contain literal bracket characters (removes brackets) before adding to the Data Model. This modifies the Excel table column headers in the worksheet. (valid for: add-to-data-mo del) |
| `--dax-query <DAXQUERY>` | DAX EVALUATE query (e.g., 'EVALUATE Sales' or 'EVALUATE SUMMARIZE(...) ') (required for: create-from-da x, update-dax) (valid for: create-from-da x, update-dax) |
| `--target-cell <TARGETCELL>` | Target cell address for table placement (default: 'A1') (valid for: create-from-da x) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
