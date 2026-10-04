### connection

Data connections (OLEDB, ODBC, ODC import)

**Actions:** `list`, `view`, `create`, `refresh`, `get-refresh-status`, `cancel-refresh`, `delete`, `load-to`, `get-properties`, `set-properties`, `test`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--connection-name <CONNECTIONNAME>` | Name of the connection to view (required for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) (valid for: view, create, refresh, get-refresh-status, cancel-refresh, delete, load-to, get-properties, set-properties, test) |
| `--connection-string <CONNECTIONSTRING>` | OLEDB or ODBC connection string (required for: create) (valid for: create, set-properties) |
| `--command-text <COMMANDTEXT>` | SQL query or table name (valid for: create, set-properties) |
| `--description <DESCRIPTION>` | Optional description for the connection (valid for: create, set-properties) |
| `--timeout <TIMEOUT>` | Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch (valid for: refresh) |
| `--sheet <SHEETNAME>` | Target worksheet name (required for: load-to) (valid for: load-to) |
| `--background-query <BACKGROUNDQUERY>` | Run query in background (null to keep current) (valid for: set-properties) |
| `--refresh-on-file-open <REFRESHONFILEOPEN>` | Refresh when file opens (null to keep current) (valid for: set-properties) |
| `--save-password <SAVEPASSWORD>` | Save password in connection (null to keep current) (valid for: set-properties) |
| `--refresh-period <REFRESHPERIOD>` | Nonnegative auto-refresh interval in minutes; 0 disables automatic refresh (null to keep current) (valid for: set-properties) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
