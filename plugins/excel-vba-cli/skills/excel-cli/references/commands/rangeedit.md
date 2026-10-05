### rangeedit

Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data

**Actions:** `advanced-filter`, `apply-filter`, `get-filters`, `clear-filters`, `remove-duplicates`, `text-to-columns`, `fill`, `auto-fill`, `create-series`, `insert-cells`, `delete-cells`, `insert-rows`, `delete-rows`, `insert-columns`, `delete-columns`, `find`, `replace`, `sort`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet containing source and output (required) |
| `--range <RANGEADDRESS>` | Complete ordinary source rectangle including headers (required for: advanced-filter, apply-filter, get-filters, clear-filters, remove-duplicates, fill, create-series, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort) (valid for: advanced-filter, apply-filter, get-filters, clear-filters, remove-duplicates, fill, create-series, insert-cells, delete-cells, insert-rows, delete-rows, insert-columns, delete-columns, find, replace, sort) |
| `--criteria-range <CRITERIARANGE>` | Native criteria rectangle including its header row (required for: advanced-filter) (valid for: advanced-filter) |
| `--mode <MODE>` | InPlace hides nonmatches; Copy writes matching records to copyToRange Accepted values (case-insensitive): InPlace, Copy. (required for: advanced-filter) (valid for: advanced-filter) |
| `--copy-to-range <COPYTORANGE>` | Copy-mode single-cell anchor or one-row selected output headers (valid for: advanced-filter) |
| `--unique-only <UNIQUEONLY>` | Retain only native unique records (valid for: advanced-filter) |
| `--overwrite-policy <OVERWRITEPOLICY>` | Copy preflight protects the full maximum output extent; allow permits authorized replacement Accepted values (case-insensitive): RejectNonempty, Allow, reject-nonempty. (valid for: advanced-filter, text-to-columns, fill, auto-fill, create-series) |
| `--column-index <COLUMNINDEX>` | One-based column relative to the rectangle (required for: apply-filter) (valid for: apply-filter) |
| `--filter-options <FILTEROPTIONS>` | Shared typed native filter options with camelCase nested keys (required for: apply-filter) (valid for: apply-filter) (JSON format) |
| `--clear-advanced <CLEARADVANCED>` | Explicit permission to clear worksheet-wide advanced row filtering when its original scope cannot be inspected; default false (valid for: clear-filters) |
| `--key-columns <KEYCOLUMNS>` | Distinct one-based columns relative to the rectangle; at least one and at most 16383 (required for: remove-duplicates) (valid for: remove-duplicates) (JSON format) |
| `--has-headers <HASHEADERS>` | Whether the first row is a header; default true, never guessed (valid for: remove-duplicates, sort) |
| `--source-range <SOURCERANGE>` | Single unmerged source column, with no header inference (required for: text-to-columns, auto-fill) (valid for: text-to-columns, auto-fill) |
| `--destination-cell <DESTINATIONCELL>` | Single unmerged output anchor; may be the source's first cell (required for: text-to-columns) (valid for: text-to-columns) |
| `--options <OPTIONS>` | Native delimiters, qualifiers, field types/positions, and number separators (required for: text-to-columns) (valid for: text-to-columns) (JSON format) |
| `--direction <DIRECTION>` | down from top row, up from bottom row, left from right column, right from left column Accepted values (case-insensitive): Down, Up, Left, Right. (required for: fill) (valid for: fill) |
| `--destination-range <DESTINATIONRANGE>` | Unmerged complete destination, including the unchanged source extent (required for: auto-fill) (valid for: auto-fill) |
| `--fill-type <FILLTYPE>` | Native AutoFill kind; default lets Excel infer the pattern Accepted values (case-insensitive): Default, Copy, Series, Formats, WithoutFormatting, Days, Weekdays, Months, Years, LinearTrend, GrowthTrend, FlashFill. (valid for: auto-fill) |
| `--orientation <ORIENTATION>` | rows across columns or columns down rows Accepted values (case-insensitive): Rows, Columns. (required for: create-series) (valid for: create-series) |
| `--series-type <SERIESTYPE>` | linear (default), growth, date, or autoFill Accepted values (case-insensitive): Linear, Growth, Date, AutoFill. (valid for: create-series) |
| `--step-value <STEPVALUE>` | Finite nonzero native step; default 1 (valid for: create-series) |
| `--stop-value <STOPVALUE>` | Optional finite stopping value; omitted fills the selected extent (valid for: create-series) |
| `--date-unit <DATEUNIT>` | Date series unit; default day Accepted values (case-insensitive): Day, Weekday, Month, Year. (valid for: create-series) |
| `--trend <TREND>` | Fit existing values for linear/growth series; requires overwritePolicy allow (valid for: create-series) |
| `--insert-shift <INSERTSHIFT>` | Direction to shift existing cells: 'Down' or 'Right' Accepted values (case-insensitive): Down, Right. (required for: insert-cells) (valid for: insert-cells) |
| `--delete-shift <DELETESHIFT>` | Direction to shift remaining cells: 'Up' or 'Left' Accepted values (case-insensitive): Up, Left. (required for: delete-cells) (valid for: delete-cells) |
| `--search-value <SEARCHVALUE>` | Text or value to search for (required for: find) (valid for: find) |
| `--find-options <FINDOPTIONS>` | Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true) (required for: find) (valid for: find) (JSON format) |
| `--max-matches <MAXMATCHES>` | Maximum matching cells to return (default: 10). Positive whole number from 1 through 2147483647. All matches are still counted. (valid for: find) |
| `--find-value <FINDVALUE>` | Text or value to search for (required for: replace) (valid for: replace) |
| `--replace-value <REPLACEVALUE>` | Text or value to replace matches with (required for: replace) (valid for: replace) |
| `--replace-options <REPLACEOPTIONS>` | Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true) (required for: replace) (valid for: replace) (JSON format) |
| `--sort-columns <SORTCOLUMNS>` | Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range (required for: sort) (valid for: sort) (JSON format) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
