### analysis

Excel what-if analysis with Goal Seek, scenarios, scenario reports, and one- or two-variable data tables

**Actions:** `goal-seek`, `list-scenarios`, `create-scenario`, `update-scenario`, `show-scenario`, `delete-scenario`, `create-scenario-summary`, `create-data-table`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet containing the what-if model (required) |
| `--formula-cell <FORMULACELL>` | Cell containing the formula whose result should reach the goal (required for: goal-seek) (valid for: goal-seek) |
| `--goal <GOAL>` | Numeric target for the formula result (required for: goal-seek) (valid for: goal-seek) |
| `--changing-cell <CHANGINGCELL>` | Single input cell Excel may adjust (required for: goal-seek) (valid for: goal-seek) |
| `--scenario-name <SCENARIONAME>` | Name of the worksheet scenario (required for: create-scenario, update-scenario, show-scenario, delete-scenario) (valid for: create-scenario, update-scenario, show-scenario, delete-scenario) |
| `--changing-cells <CHANGINGCELLS>` | Range of input cells whose values the scenario stores (required for: create-scenario, update-scenario) (valid for: create-scenario, update-scenario) |
| `--values <VALUES>` | One value per changing cell, in range order (required for: create-scenario, update-scenario) (valid for: create-scenario, update-scenario) (JSON format) |
| `--comment <COMMENT>` | Optional scenario description (valid for: create-scenario) |
| `--locked <LOCKED>` | Prevent scenario editing when worksheet protection is enabled (valid for: create-scenario) |
| `--hidden <HIDDEN>` | Hide the scenario when worksheet protection is enabled (valid for: create-scenario) |
| `--report-type <REPORTTYPE>` | Scenario report type: Summary or PivotTable Accepted values (case-insensitive): Summary, PivotTable. (valid for: create-scenario-summary) |
| `--result-cells <RESULTCELLS>` | Formula result cells to include in the scenario report (valid for: create-scenario-summary) |
| `--table-range <TABLERANGE>` | Prepared sensitivity table range, including formulas and trial input values (required for: create-data-table) (valid for: create-data-table) |
| `--row-input-cell <ROWINPUTCELL>` | Model input cell to substitute values from the table's row; supply at least one input cell (valid for: create-data-table) |
| `--column-input-cell <COLUMNINPUTCELL>` | Model input cell to substitute values from the table's column; supply at least one input cell (valid for: create-data-table) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
