### range

Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions

**Actions:** `trace-precedents`, `trace-dependents`, `get-values`, `set-values`, `get-formulas`, `get-spill-info`, `set-formulas`, `validate-formulas`, `clear-all`, `clear-contents`, `clear-formats`, `copy`, `get-number-formats`, `set-number-format`, `set-number-formats`, `get-used-range`, `get-current-region`, `get-info`, `get-special-cells`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet name; empty for a named range (required for: trace-precedents, trace-dependents, get-values, set-values, get-formulas, get-spill-info, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info, get-special-cells) (valid for: trace-precedents, trace-dependents, get-values, set-values, get-formulas, get-spill-info, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-used-range, get-current-region, get-info, get-special-cells) |
| `--range <RANGEADDRESS>` | Exact starting cells, rectangle, disjoint areas, or named range (required for: trace-precedents, trace-dependents, get-values, set-values, get-formulas, get-spill-info, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info, get-special-cells) (valid for: trace-precedents, trace-dependents, get-values, set-values, get-formulas, get-spill-info, set-formulas, validate-formulas, clear-all, clear-contents, clear-formats, get-number-formats, set-number-format, set-number-formats, get-info, get-special-cells) |
| `--values <VALUES>` | 2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Strict ISO dates such as "2025-01-15" become native Excel dates. Optional if valuesFile is provided. (valid for: set-values) (JSON format) |
| `--values-file <VALUESFILE>` | Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter. (valid for: set-values) |
| `--overwrite-policy <OVERWRITEPOLICY>` | reject-nonempty (default) checks all direct destinations and rejects existing content, including formulas displaying blank. allow permits intentional replacement, not bypassing Excel protection. Inspection failure stops the operation; no rollback or interactive-edit isolation. Accepted values (case-insensitive): RejectNonempty, Allow, reject-nonempty. (valid for: set-values, set-formulas, copy) |
| `--reference-style <REFERENCESTYLE>` | a1 (default) or r1c1 native formula notation; range addresses remain A1 Accepted values (case-insensitive): A1, R1C1. (valid for: get-formulas, set-formulas) |
| `--formulas <FORMULAS>` | 2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided. (valid for: set-formulas, validate-formulas) (JSON format) |
| `--formulas-file <FORMULASFILE>` | Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter. (valid for: set-formulas, validate-formulas) |
| `--source-sheet <SOURCESHEET>` | Source worksheet name for copy operations (required for: copy) (valid for: copy) |
| `--source-range <SOURCERANGE>` | Source range address for copy operations (e.g., 'A1:D10') (required for: copy) (valid for: copy) |
| `--target-sheet <TARGETSHEET>` | Target worksheet name for copy operations (required for: copy) (valid for: copy) |
| `--target-range <TARGETRANGE>` | Target range address - can be single cell for paste destination (e.g., 'A1') (required for: copy) (valid for: copy) |
| `--paste-kind <PASTEKIND>` | Required native paste kind: all, values, formulas, formats, or validation Accepted values (case-insensitive): All, Values, Formulas, Formats, Validation. (required for: copy) (valid for: copy) |
| `--transpose <TRANSPOSE>` | Exchange source rows and columns; validation uses transposed destination dimensions (valid for: copy) |
| `--skip-blanks <SKIPBLANKS>` | Preserve destinations corresponding to native blank source cells; formulas displaying blank are not skipped (valid for: copy) |
| `--format-code <FORMATCODE>` | Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text) (required for: set-number-format) (valid for: set-number-format) |
| `--formats <FORMATS>` | 2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided. (valid for: set-number-formats) (JSON format) |
| `--formats-file <FORMATSFILE>` | Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter. (valid for: set-number-formats) |
| `--cell-address <CELLADDRESS>` | Single cell address (e.g., 'B5') - expands to contiguous data region around this cell (required for: get-current-region) (valid for: get-current-region) |
| `--cell-kind <CELLKIND>` | Cell selector: formulas, constants, blanks, errors, or visible Accepted values (case-insensitive): Formulas, Constants, Blanks, Errors, Visible. (required for: get-special-cells) (valid for: get-special-cells) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
