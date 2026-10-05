### rangelink

Hyperlink, threaded comment, and cell protection operations for Excel ranges

**Actions:** `add-hyperlink`, `update-hyperlink`, `remove-hyperlink`, `list-hyperlinks`, `get-hyperlink`, `add-threaded-comment`, `list-threaded-comments`, `add-threaded-comment-reply`, `delete-threaded-comment`, `set-cell-protection`, `get-cell-protection`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Name of the worksheet (required) |
| `--cell-address <CELLADDRESS>` | Single cell address (e.g., 'A1') (required for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) (valid for: add-hyperlink, update-hyperlink, get-hyperlink, add-threaded-comment, list-threaded-comments, add-threaded-comment-reply, delete-threaded-comment) |
| `--url <URL>` | Optional external URL or file path. Omit for an internal workbook link. (valid for: add-hyperlink, update-hyperlink) |
| `--display-text <DISPLAYTEXT>` | Text to display in the cell (optional, defaults to URL) (valid for: add-hyperlink, update-hyperlink) |
| `--tooltip <TOOLTIP>` | Tooltip text shown on hover (optional) (valid for: add-hyperlink, update-hyperlink) |
| `--sub-address <SUBADDRESS>` | Optional internal workbook target such as "'Sheet2'!A1" (valid for: add-hyperlink, update-hyperlink) |
| `--range <RANGEADDRESS>` | Cell range address to remove hyperlinks from (e.g., 'A1:D10') (required for: remove-hyperlink, set-cell-protection, get-cell-protection) (valid for: remove-hyperlink, set-cell-protection, get-cell-protection) |
| `--text <TEXT>` | Comment or reply text; cloud mentions and assignments are not supported (required for: add-threaded-comment, add-threaded-comment-reply) (valid for: add-threaded-comment, add-threaded-comment-reply) |
| `--locked <LOCKED>` | Optional native lock flag; omitted leaves it unchanged (valid for: set-cell-protection) |
| `--formula-hidden <FORMULAHIDDEN>` | Optional native formula-hiding flag; omitted leaves it unchanged (valid for: set-cell-protection) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
