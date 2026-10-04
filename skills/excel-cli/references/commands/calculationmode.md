### calculationmode

Read/change application calculation settings and explicit workbook precision

**Actions:** `get-settings`, `set-settings`, `set-precision`, `calculate`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--mode <MODE>` | Optional automatic/manual/semi -automatic mode Accepted values (case-insensitive): Automatic, Manual, SemiAutomatic. (valid for: set-settings) |
| `--iteration-enabled <ITERATIONENABLED>` | Optional circular-formula iteration flag (valid for: set-settings) |
| `--maximum-iterations <MAXIMUMITERATIONS>` | Optional iteration count, 1 through 32767 (valid for: set-settings) |
| `--maximum-change <MAXIMUMCHANGE>` | Optional finite positive convergence tolerance (valid for: set-settings) |
| `--calculate-before-save <CALCULATEBEFORESAVE>` | Optional native calculate-before-save flag (valid for: set-settings) |
| `--precision-as-displayed <PRECISIONASDISPLAYED>` | Required true to enable or false to disable (required for: set-precision) (valid for: set-precision) |
| `--allow-precision-loss <ALLOWPRECISIONLOSS>` | Required explicit true permission when enabling; default false (valid for: set-precision) |
| `--scope <SCOPE>` | Required application, sheet, or range; former workbook spelling is removed Accepted values (case-insensitive): Application, Sheet, Range. (required for: calculate) (valid for: calculate) |
| `--sheet <SHEETNAME>` | Required for sheet/range scope (valid for: calculate) |
| `--range <RANGEADDRESS>` | Required for range scope (valid for: calculate) |
| `--kind <KIND>` | normal, full, or rebuild; default normal Accepted values (case-insensitive): Normal, Full, Rebuild. (valid for: calculate) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
