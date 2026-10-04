### worksheetstyle

Worksheet styling, visibility, protection, grouping, and outline operations

**Actions:** `set-tab-color`, `get-tab-color`, `clear-tab-color`, `set-protection`, `get-protection`, `set-comment`, `get-comment`, `clear-comment`, `add-image`, `get-image-count`, `add-shape`, `get-shape-count`, `set-page-setup`, `get-page-setup`, `get-page-breaks`, `set-page-breaks`, `set-visibility`, `get-visibility`, `show`, `hide`, `very-hide`, `group`, `ungroup`, `get-outline-info`, `set-outline-settings`, `show-outline-levels`, `clear-outline`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Name of the worksheet to color (required) |
| `--red <RED>` | Red color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--green <GREEN>` | Green color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--blue <BLUE>` | Blue color component (0-255) (required for: set-tab-color) (valid for: set-tab-color) |
| `--is-protected <ISPROTECTED>` | Whether the worksheet should be protected (required for: set-protection) (valid for: set-protection) |
| `--password <PASSWORD>` | Optional password for protecting/unprotecting the sheet (valid for: set-protection) |
| `--options <OPTIONS>` | Optional native protection permissions; valid only when protecting. Nested JSON uses camelCase. (valid for: set-protection) |
| `--cell-address <CELLADDRESS>` | Cell address such as A1 (required for: set-comment, get-comment, clear-comment, add-image, add-shape) (valid for: set-comment, get-comment, clear-comment, add-image, add-shape) |
| `--text <TEXT>` | Cell note text to set (required for: set-comment) (valid for: set-comment) |
| `--image-path <IMAGEPATH>` | Absolute path to the image file on disk (required for: add-image) (valid for: add-image) |
| `--orientation <ORIENTATION>` | Optional page orientation: 'portrait' or 'landscape'; omitted leaves unchanged (valid for: set-page-setup) |
| `--fit-to-pages-wide <FITTOPAGESWIDE>` | Number of pages wide, or zero for unlimited; selects fit mode (valid for: set-page-setup) |
| `--fit-to-pages-tall <FITTOPAGESTALL>` | Number of pages tall, or zero for unlimited; selects fit mode (valid for: set-page-setup) |
| `--center-horizontally <CENTERHORIZONTALLY>` | Whether to center the printout horizontally on the page (valid for: set-page-setup) |
| `--center-vertically <CENTERVERTICALLY>` | Whether to center the printout vertically on the page (valid for: set-page-setup) |
| `--page-setup-options <PAGESETUPOPTIONS>` | Native print scope, titles, point margins, headers/footers, paper, order and scaling. Nested camelCase; null preserves, empty text clears. Zoom conflicts with fit settings. Does not print or open preview. (valid for: set-page-setup) (JSON format) |
| `--page-break-options <PAGEBREAKOPTIONS>` | Required rows and columns lists with one-based positions before which to break. Both lists replace existing manual breaks, including those outside the current print scope. (required for: set-page-breaks) (valid for: set-page-breaks) (JSON format) |
| `--visibility <VISIBILITY>` | Visibility level: 'visible', 'hidden', or 'veryhidden' Accepted values (case-insensitive): Visible, Hidden, VeryHidden. (required for: set-visibility) (valid for: set-visibility) |
| `--range <RANGEADDRESS>` | Row or column range to group (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--axis <AXIS>` | Grouping axis: Rows or Columns Accepted values (case-insensitive): Rows, Columns. (required for: group, ungroup, get-outline-info) (valid for: group, ungroup, get-outline-info) |
| `--summary-row <SUMMARYROW>` | Summary row position: above or below (valid for: set-outline-settings) |
| `--summary-column <SUMMARYCOLUMN>` | Summary column position: left or right (valid for: set-outline-settings) |
| `--automatic-styles <AUTOMATICSTYLES>` | Whether Excel applies automatic outline styles (valid for: set-outline-settings) |
| `--row-levels <ROWLEVELS>` | Optional row outline level to display (valid for: show-outline-levels) |
| `--column-levels <COLUMNLEVELS>` | Optional column outline level to display (valid for: show-outline-levels) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
