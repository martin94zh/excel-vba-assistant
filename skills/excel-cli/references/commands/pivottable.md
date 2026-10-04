### pivottable

PivotTable lifecycle management: create from various sources, list, read details, refresh, and delete

**Actions:** `list`, `read`, `create-from-range`, `create-from-table`, `create-from-datamodel`, `delete`, `refresh`, `get-source`, `set-source`, `get-cache-options`, `set-cache-options`, `drill-through`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--pivot-table-name <PIVOTTABLENAME>` | Name of the PivotTable (required for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-source, set-source, get-cache-options, set-cache-options, drill-through) (valid for: read, create-from-range, create-from-table, create-from-datamodel, delete, refresh, get-source, set-source, get-cache-options, set-cache-options, drill-through) |
| `--source-sheet <SOURCESHEET>` | Source worksheet name (required for: create-from-range) (valid for: create-from-range) |
| `--source-range <SOURCERANGE>` | Source range address (e.g., "A1:F100") (required for: create-from-range) (valid for: create-from-range) |
| `--destination-sheet <DESTINATIONSHEET>` | Destination worksheet name (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--destination-cell <DESTINATIONCELL>` | Destination cell address (e.g., "A1") (required for: create-from-range, create-from-table, create-from-datamodel) (valid for: create-from-range, create-from-table, create-from-datamodel) |
| `--table-name <TABLENAME>` | Name of the Excel Table (required for: create-from-table, create-from-datamodel) (valid for: create-from-table, create-from-datamodel, set-source) |
| `--timeout <TIMEOUT>` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--source-sheet-name <SOURCESHEETNAME>` | New source worksheet. (required for: set-source) (valid for: set-source) |
| `--source-range-address <SOURCERANGEADDRESS>` | One contiguous header/data range with the same field names; mutually exclusive with tableName. (valid for: set-source) |
| `--enable-refresh <ENABLEREFRESH>` | Allow the PivotCache to refresh (valid for: set-cache-options) |
| `--refresh-on-file-open <REFRESHONFILEOPEN>` | Refresh the PivotCache when the workbook opens (valid for: set-cache-options) |
| `--missing-items-limit <MISSINGITEMSLIMIT>` | How many deleted source items Excel retains in the cache Accepted values (case-insensitive): Default, None, Max, Max2. (valid for: set-cache-options) |
| `--optimize-cache <OPTIMIZECACHE>` | Optimize the cache when it is constructed (valid for: set-cache-options) |
| `--save-source-data <SAVESOURCEDATA>` | Save source records with the PivotTable (valid for: set-cache-options) |
| `--cell-address <CELLADDRESS>` | Value-cell address on the PivotTable worksheet (for example, G4) (required for: drill-through) (valid for: drill-through) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
