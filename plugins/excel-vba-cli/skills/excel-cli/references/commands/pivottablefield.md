### pivottablefield

PivotTable field management: add/remove/configure fields, filtering, sorting, and grouping

**Actions:** `list-fields`, `add-row-field`, `add-column-field`, `add-value-field`, `add-filter-field`, `remove-field`, `set-field-function`, `set-field-calculation`, `set-field-name`, `set-field-format`, `set-field-filter`, `get-field-filters`, `add-field-filter`, `clear-field-filters`, `get-item-expansion`, `set-item-expansion`, `sort-field`, `group-by-date`, `group-by-numeric`, `group-items`, `ungroup-field`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--pivot-table-name <PIVOTTABLENAME>` | Name of the PivotTable (required) |
| `--field-name <FIELDNAME>` | Name of the field to add (required for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-calculation, set-field-name, set-field-format, set-field-filter, get-field-filters, add-field-filter, clear-field-filters, get-item-expansion, set-item-expansion, sort-field, group-by-date, group-by-numeric, group-items) (valid for: add-row-field, add-column-field, add-value-field, add-filter-field, remove-field, set-field-function, set-field-calculation, set-field-name, set-field-format, set-field-filter, get-field-filters, add-field-filter, clear-field-filters, get-item-expansion, set-item-expansion, sort-field, group-by-date, group-by-numeric, group-items) |
| `--position <POSITION>` | Optional position in row area (1-based) (valid for: add-row-field, add-column-field) |
| `--aggregation-function <AGGREGATIONFUNCTION>` | Aggregation function (for Regular and OLAP auto-create mode) Accepted values (case-insensitive): Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP. (required for: set-field-function) (valid for: add-value-field, set-field-function) |
| `--custom-name <CUSTOMNAME>` | Optional custom name for the field/measure (required for: set-field-name) (valid for: add-value-field, set-field-name) |
| `--calculation <CALCULATION>` | Native additional calculation; Normal preserves aggregation Accepted values (case-insensitive): Normal, DifferenceFrom, PercentOf, PercentDifferenceFrom, RunningTotal, PercentOfRow, PercentOfColumn, PercentOfTotal, Index, PercentOfParentRow, PercentOfParentColumn, PercentOfParent, PercentRunningTotal, RankAscending, RankDescending. (required for: set-field-calculation) (valid for: set-field-calculation) |
| `--base-field-name <BASEFIELDNAME>` | Required exact row/column field for differences, running totals, parent-field percentage, and ranks; otherwise omit (valid for: set-field-calculation) |
| `--base-item-kind <BASEITEMKIND>` | Required for DifferenceFrom, PercentOf, PercentDifferenceFrom: Named, Previous, or Next; otherwise omit Accepted values (case-insensitive): Named, Previous, Next. (valid for: set-field-calculation) |
| `--base-item-name <BASEITEMNAME>` | Exact native item name, required only with Named; never a numeric index (valid for: set-field-calculation) |
| `--number-format <NUMBERFORMAT>` | Number format string (required for: set-field-format) (valid for: set-field-format) |
| `--selected-values <SELECTEDVALUES>` | Values to show (others will be hidden) (required for: set-field-filter) (valid for: set-field-filter) (JSON format) |
| `--filter-options <FILTEROPTIONS>` | Typed filter type and matching text, number, or date criteria. Value/top filters require an exact dataFieldName caption. (required for: add-field-filter) (valid for: add-field-filter) (JSON format) |
| `--item-name <ITEMNAME>` | Exact item caption. (required for: get-item-expansion, set-item-expansion) (valid for: get-item-expansion, set-item-expansion) |
| `--expanded <EXPANDED>` | True expands; false collapses. (required for: set-item-expansion) (valid for: set-item-expansion) |
| `--direction <DIRECTION>` | Sort direction Accepted values (case-insensitive): Ascending, Descending. (valid for: sort-field) |
| `--interval <INTERVAL>` | Grouping interval (Months, Quarters, Years) Accepted values (case-insensitive): Days, Months, Quarters, Years. (required for: group-by-date) (valid for: group-by-date) |
| `--start <START>` | Starting value (null = use field minimum) (valid for: group-by-numeric) |
| `--end-value <ENDVALUE>` | Ending value (null = use field maximum) (valid for: group-by-numeric) |
| `--interval-size <INTERVALSIZE>` | Size of each group (e.g., 100 for groups of 100) (required for: group-by-numeric) (valid for: group-by-numeric) |
| `--item-names <ITEMNAMES>` | Exact item captions to group (required for: group-items) (valid for: group-items) (JSON format) |
| `--group-name <GROUPNAME>` | Caption for the new group (required for: group-items) (valid for: group-items) |
| `--grouped-field-name <GROUPEDFIELDNAME>` | Generated grouped field name (required for: ungroup-field) (valid for: ungroup-field) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
