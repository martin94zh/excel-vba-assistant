### rangeformat

Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions

**Actions:** `get-visibility`, `set-visibility`, `get-format`, `set-style`, `get-style`, `format`, `validate-range`, `get-validation`, `remove-validation`, `auto-fit-columns`, `auto-fit-rows`, `merge-cells`, `unmerge-cells`, `get-merge-info`, `set-column-width`, `set-row-height`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet name; empty for named ranges (required) |
| `--range <RANGEADDRESS>` | Exact scope; whole intersecting dimensions are returned without a cap (required for: get-visibility, set-visibility, get-format, set-style, get-style, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) (valid for: get-visibility, set-visibility, get-format, set-style, get-style, validate-range, get-validation, remove-validation, auto-fit-columns, auto-fit-rows, merge-cells, unmerge-cells, get-merge-info, set-column-width, set-row-height) |
| `--axis <AXIS>` | rows or columns Accepted values (case-insensitive): Rows, Columns. (required for: get-visibility, set-visibility) (valid for: get-visibility, set-visibility) |
| `--hidden <HIDDEN>` | Required true to hide or false to show; native stored dimensions are retained (required for: set-visibility) (valid for: set-visibility) |
| `--view <VIEW>` | stored (default), displayed, or both formatting snapshots Accepted values (case-insensitive): Stored, Displayed, Both. (valid for: get-format) |
| `--style-name <STYLENAME>` | Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset. (required for: set-style) (valid for: set-style) |
| `--range-addresses <RANGEADDRESSES>` | One or more target range addresses; all are validated before writing (required for: format) (valid for: format) |
| `--format-options <FORMATOPTIONS>` | Typed JSON object: fontName, fontSize, bold, italic, underline (none/single/double/single Accounting/doubleAccountin g), strikethrough, subscript, superscript, themeFont (0 none/1 major/2 minor), fontColor/fontThemeColor/f ontTintAndShade, fillColor/fillThemeColor/f illTintAndShade, borders (position, lineStyle, weight, color/themeColor/tintAndSh ade), horizontalAlignment, verticalAlignment, wrapText, shrinkToFit, indentLevel (0-15), readingOrder (context/leftToRight/right ToLeft), orientation, numberFormat. Theme-color indices 1-12; tints -1 to 1. Border positions Left/Top/Bottom/Right/Insi deHorizontal/InsideVertica l/DiagonalUp/DiagonalDown. Nested keys remain camelCase. (required for: format) (valid for: format) |
| `--validation-type <VALIDATIONTYPE>` | Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom' (required for: validate-range) (valid for: validate-range) |
| `--validation-operator <VALIDATIONOPERATOR>` | Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual' (valid for: validate-range) |
| `--formula1 <FORMULA1>` | First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"' (valid for: validate-range) |
| `--formula2 <FORMULA2>` | Second validation formula/value - required only for 'between' and 'notBetween' operators (valid for: validate-range) |
| `--show-input-message <SHOWINPUTMESSAGE>` | Whether to show input message when cell is selected (default: false) (valid for: validate-range) |
| `--input-title <INPUTTITLE>` | Title for the input message popup (valid for: validate-range) |
| `--input-message <INPUTMESSAGE>` | Text for the input message popup (valid for: validate-range) |
| `--show-error-alert <SHOWERRORALERT>` | Whether to show error alert on invalid input (default: true) (valid for: validate-range) |
| `--error-style <ERRORSTYLE>` | Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry) (valid for: validate-range) |
| `--error-title <ERRORTITLE>` | Title for the error alert popup (valid for: validate-range) |
| `--error-message <ERRORMESSAGE>` | Text for the error alert popup (valid for: validate-range) |
| `--ignore-blank <IGNOREBLANK>` | Whether to allow blank cells in validation (default: true) (valid for: validate-range) |
| `--show-dropdown <SHOWDROPDOWN>` | Whether to show dropdown arrow for list validation (default: true) (valid for: validate-range) |
| `--column-width <COLUMNWIDTH>` | Width in Excel character-width units, not points. Standard width is approximately 8.43. Range: 0.25-409. (required for: set-column-width) (valid for: set-column-width) |
| `--row-height <ROWHEIGHT>` | Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points. (required for: set-row-height) (valid for: set-row-height) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
