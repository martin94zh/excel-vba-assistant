### powerquery

Power Query M code and data loading

**Actions:** `list`, `view`, `refresh`, `get-load-config`, `delete`, `create`, `update`, `load-to`, `refresh-all`, `rename`, `unload`, `evaluate`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--query-name <QUERYNAME>` | Name of the query to view (required for: view, refresh, get-load-config, delete, create, update, load-to, unload) (valid for: view, refresh, get-load-config, delete, create, update, load-to, unload) |
| `--timeout <TIMEOUT>` | Public input is whole seconds from 0 through 2147483. Omitted or 0 uses the 30-minute data-operation default. (valid for: refresh, refresh-all) |
| `--m-code <MCODE>` | Raw M code. Public callers must supply either inline mCode or a readable mCodeFile, not both. (required for: create, update, evaluate) (valid for: create, update, evaluate) |
| `--m-code-file <MCODEFILE>` | Path to a readable file containing mCode; use instead of inline mCode, not together (valid for: create, update, evaluate) |
| `--load-destination <LOADDESTINATION>` | Load destination mode Accepted values (case-insensitive): ConnectionOnly, LoadToTable, LoadToDataModel, LoadToBoth, worksheet, table, data-model, datamodel, both. (required for: load-to) (valid for: create, load-to) |
| `--target-sheet <TARGETSHEET>` | Target worksheet name (required for LoadToTable and LoadToBoth; defaults to query name when omitted) (valid for: create, load-to) |
| `--target-cell-address <TARGETCELLADDRESS>` | Optional target cell address for worksheet loads (e.g., "B5"). Required when loading to an existing worksheet with other data. (valid for: create, load-to) |
| `--format-m-code <FORMATMCODE>` | Whether to send M code to the remote powerqueryformatter.com service before saving. Defaults to false to preserve privacy. (valid for: create, update) |
| `--refresh <REFRESH>` | Whether to refresh data after update (default: true) (valid for: update) |
| `--old-name <OLDNAME>` | Current name of the query (required for: rename) (valid for: rename) |
| `--new-name <NEWNAME>` | New name for the query (required for: rename) (valid for: rename) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
