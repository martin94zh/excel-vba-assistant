### window

Control Excel window visibility, position, state, status bar, and worksheet-specific views

**Actions:** `get-context`, `show`, `hide`, `bring-to-front`, `get-info`, `set-state`, `set-position`, `arrange`, `set-status-bar`, `clear-status-bar`, `get-view`, `freeze-panes`, `unfreeze-panes`, `set-split`, `set-zoom`, `set-display-options`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--window-state <WINDOWSTATE>` | Window state: 'normal', 'minimized', or 'maximized' (required for: set-state) (valid for: set-state) |
| `--left <LEFT>` | Window left position in points (valid for: set-position) |
| `--top <TOP>` | Window top position in points (valid for: set-position) |
| `--width <WIDTH>` | Window width in points (valid for: set-position) |
| `--height <HEIGHT>` | Window height in points (valid for: set-position) |
| `--preset <PRESET>` | Preset name: 'left-half', 'right-half', 'top-half', 'bottom-half', 'center', 'full-screen' (required for: arrange) (valid for: arrange) |
| `--text <TEXT>` | Status bar text to display (e.g. "Building PivotTable from Sales data...") (required for: set-status-bar) (valid for: set-status-bar) |
| `--sheet <SHEETNAME>` | Worksheet whose view should be inspected (required for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) (valid for: get-view, freeze-panes, unfreeze-panes, set-split, set-zoom, set-display-options) |
| `--frozen-rows <FROZENROWS>` | Number of rows to freeze from the top (0-1,048,575) (valid for: freeze-panes) |
| `--frozen-columns <FROZENCOLUMNS>` | Number of columns to freeze from the left (0-16,383) (valid for: freeze-panes) |
| `--split-rows <SPLITROWS>` | Number of rows above the horizontal split (0-1,048,575) (valid for: set-split) |
| `--split-columns <SPLITCOLUMNS>` | Number of columns left of the vertical split (0-16,383) (valid for: set-split) |
| `--zoom <ZOOM>` | Zoom percentage from 10 through 400 (required for: set-zoom) (valid for: set-zoom) |
| `--show-gridlines <SHOWGRIDLINES>` | Whether to display cell gridlines (valid for: set-display-options) |
| `--show-headings <SHOWHEADINGS>` | Whether to display row and column headings (valid for: set-display-options) |
| `--show-outline-symbols <SHOWOUTLINESYMBOLS>` | Whether to display outline level symbols (valid for: set-display-options) |
| `--show-formulas <SHOWFORMULAS>` | Whether to display formulas instead of their calculated values (valid for: set-display-options) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
