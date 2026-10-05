### pivottablecalc

PivotTable calculated fields/members, layout configuration, and data extraction

**Actions:** `get-data`, `create-calculated-field`, `list-calculated-fields`, `delete-calculated-field`, `list-calculated-members`, `create-calculated-member`, `delete-calculated-member`, `set-layout`, `get-layout-options`, `set-layout-options`, `set-subtotals`, `set-grand-totals`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--pivot-table-name <PIVOTTABLENAME>` | Name of the PivotTable (required) |
| `--field-name <FIELDNAME>` | Name for the calculated field (required for: create-calculated- field, delete-calculated- field, set-subtotals) (valid for: create-calculated- field, delete-calculated- field, set-subtotals) |
| `--formula <FORMULA>` | Formula using field references (e.g., "=Revenue-Cost") (required for: create-calculated- field, create-calculated- member) (valid for: create-calculated- field, create-calculated- member) |
| `--member-name <MEMBERNAME>` | Name for the calculated member (MDX naming format) (required for: create-calculated- member, delete-calculated- member) (valid for: create-calculated- member, delete-calculated- member) |
| `--type <TYPE>` | Type of calculated member (Member, Set, or Measure) Accepted values (case-insensitive) : Member, Set, Measure. (valid for: create-calculated- member) |
| `--solve-order <SOLVEORDER>` | Solve order for calculation precedence (default: 0) (valid for: create-calculated- member) |
| `--display-folder <DISPLAYFOLDER>` | Display folder path for organizing measures (optional) (valid for: create-calculated- member) |
| `--number-format <NUMBERFORMAT>` | Number format code for the calculated member (optional) (valid for: create-calculated- member) |
| `--row-layout <ROWLAYOUT>` | Layout form: 0=Compact, 1=Tabular, 2=Outline (required for: set-layout) (valid for: set-layout) |
| `--layout-options <LAYOUTOPTIONS>` | Typed row layout, repeated labels, style name, preserveFormatting , and header/banding options. (required for: set-layout-options ) (valid for: set-layout-options ) (JSON format) |
| `--show-subtotals <SHOWSUBTOTALS>` | True to show automatic subtotals, false to hide (required for: set-subtotals) (valid for: set-subtotals) |
| `--show-row-grand-totals <SHOWROWGRANDTOTALS>` | Show row grand totals (bottom summary row) (required for: set-grand-totals) (valid for: set-grand-totals) |
| `--show-column-grand-totals <SHOWCOLUMNGRANDTOTALS>` | Show column grand totals (right summary column) (required for: set-grand-totals) (valid for: set-grand-totals) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
