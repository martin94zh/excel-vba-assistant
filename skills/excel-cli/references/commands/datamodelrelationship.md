### datamodelrelationship

Data Model relationships - link tables for cross-table DAX calculations

**Actions:** `list-relationships`, `read-relationship`, `create-relationship`, `update-relationship`, `delete-relationship`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--from-table <FROMTABLE>` | Source table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--from-column <FROMCOLUMN>` | Source column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-table <TOTABLE>` | Target table name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--to-column <TOCOLUMN>` | Target column name (required for: read-relationship, create-relationship, update-relationship, delete-relationship) (valid for: read-relationship, create-relationship, update-relationship, delete-relationship) |
| `--active <ACTIVE>` | Whether the relationship should be active (default: true) (required for: update-relationship) (valid for: create-relationship, update-relationship) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
