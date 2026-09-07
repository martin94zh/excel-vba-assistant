---
name: excel-vba-assistant
description: >
  Excel VBA Assistant plugin skill for Windows workbook automation via the embedded `excel-mcp` MCP server.
  The plugin automatically opens Excel through MCP when the user selects a workbook; AI can then call
  MCP tools (range, table, pivot, powerquery, datamodel, vba, etc.) directly against that session.
  Use when the task involves Excel workbooks, worksheets, ranges, tables, PivotTables, charts,
  Power Query (M code), Data Model/DAX, VBA macros, slicers, conditional formatting, named ranges,
  connections, calculation mode, screenshots, or window management.
  Triggers: Excel, spreadsheet, workbook, xlsx, xlsm, Power Query, DAX, PivotTable, chart, dashboard, VBA, MCP.
---

# Excel VBA Assistant MCP

Use the `excel-mcp` MCP server when the task involves **Excel workbooks, VBA code, worksheets, ranges, tables, PivotTables, charts, Power Query, DAX, or window management**.

This plugin embeds `sbroenne/mcp-server-excel` and exposes **16 tool categories with 230+ operations**. When the user selects a workbook, the plugin automatically opens it through MCP and creates a session; AI can then call `range`, `table`, `pivottable`, `powerquery`, `datamodel`, `vba`, and other MCP tools directly. Any operation not yet wrapped can be invoked directly via `callTool`.

## Preconditions

- Windows host with Microsoft Excel installed (2016+).
- Use full Windows paths: `C:\Users\Name\Documents\Report.xlsx`.
- The plugin automatically opens the selected workbook through MCP; **do not tell the user to manually open Excel first**. If the file is already open in another Excel instance, the plugin detects the conflict and offers to take over.
- Excel must have **"Trust access to the VBA project object model"** enabled, or all VBA read/write operations fail.
- The plugin does not automatically detect Excel closure. The user must click **断开 Excel** in the control panel to clean up the MCP session.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Ensure session | `file` | `open` with path, or verify via `list` | Always first |
| 2. Create sheets | `worksheet` | `create`, `rename` | If needed |
| 3. Write data | `range` | `set-values` | Always (2D arrays) |
| 4. Format | `range` / `range_format` | `set-number-format`, `format-range` | After writing |
| 5. Structure | `table` | `create` | Convert data to tables |
| 6. Save & close | `file` | `close` with `save: true` | Always last |

**Important**: When the user selects a workbook in the plugin (e.g. via `Excel VBA: 选择 Excel 文件`), the plugin automatically opens it through MCP and creates a session. AI can then call MCP tools directly using that session. If no session exists yet, use `file(open, path)` with the known file path. Do NOT ask the user to manually open Excel.

## Calculation Mode Workflow (Batch Performance)

Use `calculation_mode` for **bulk write performance optimization**. When writing many values or formulas, disable auto-recalc to avoid recalculating after every cell:

```
1. calculation_mode(action: 'set-mode', mode: 'manual')   → Disable auto-recalc
2. Perform all writes (range set-values, set-formulas)
3. calculation_mode(action: 'calculate', scope: 'workbook') → Recalculate once
4. calculation_mode(action: 'set-mode', mode: 'automatic')  → Restore default
```

**Note:** You do NOT need manual mode to read formulas — `range(get-formulas)` returns formula text regardless of calculation mode.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: NEVER Ask Clarifying Questions

**STOP.** If you're about to ask "Which file?", "What table?", "Where should I put this?" — DON'T.

| Bad (Asking) | Good (Discovering) |
|--------------|-------------------|
| "Which Excel file should I use?" | `file(list)` → use the open session |
| "What's the table name?" | `table(list)` → discover tables |
| "Which sheet has the data?" | `worksheet(list)` → check all sheets |
| "Should I create a PivotTable?" | YES — create it on a new sheet |

**You have tools to answer your own questions. USE THEM.**

Only ask when the request is genuinely ambiguous (e.g., "update the data" without specifying what data or which file).

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a tool call.** After completing all operations, always provide a brief text message confirming what was done. Silent tool-call-only responses are incomplete.

### Rule 3: Format Data Professionally

Always apply number formats after setting values:

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |

**Workflow:**

```
1. range(action: 'set-values', ...)          → data is now in cells
2. range(action: 'set-number-format', ...)   → apply format
```

### Rule 4: Use Excel Tables (Not Plain Ranges)

Always convert tabular data to Excel Tables:

```
1. range(action: 'set-values', ...)          → write data including headers
2. table(action: 'create', table_name: 'SalesData', range_address: 'A1:D100')
```

**Why:** Structured references, auto-expand, required for Data Model/DAX.

### Rule 5: Session Lifecycle

The plugin manages the Excel session lifecycle: it opens Excel through MCP when the user selects a workbook, caches the `session_id`, and closes the session when the user clicks **断开 Excel**. AI operations still require a valid `session_id`.

```
1. Plugin opens workbook via MCP  → session_id (cached internally)
2. AI calls tools using session_id
3. Plugin closes workbook via MCP on disconnect
```

If you open a session yourself with `file(action: 'open', path: '...')`, always close it with `file(action: 'close', session_id: '...', save: true)` when done. **Unclosed sessions leave Excel processes running, locking files.**

### Rule 6: Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: table(create, ...)                        → Table exists
Step 2: table(action: 'add-to-datamodel', ...)    → Table in Data Model
Step 3: datamodel(action: 'create-measure', ...)  → NOW this works
```

### Rule 7: Power Query Development Lifecycle

**BEST PRACTICE: Test-First Workflow**

```
1. powerquery(action: 'evaluate', m_code: '...')  → Test WITHOUT persisting
2. powerquery(action: 'create', ...)              → Store validated query
3. powerquery(action: 'refresh', ...)             → Load data
```

**Why evaluate first:**
- Catches syntax errors and missing sources BEFORE creating permanent queries.
- Better error messages than COM exceptions from create/update.
- See actual data preview (columns + sample rows).
- No cleanup needed — like a REPL for M code.
- Skip only for trivial literal tables.

**Common mistake:** Creating/updating without evaluate pollutes the workbook with broken queries.

### Rule 8: Targeted Updates Over Delete-Rebuild

- **Prefer**: `range(set-values)` on a specific range (e.g., `A5:C5` for row 5).
- **Avoid**: Deleting and recreating entire structures.

**Why:** Preserves formatting, formulas, and references.

### Rule 9: Follow suggestedNextActions

Error responses include actionable hints:

```json
{
  "success": false,
  "errorMessage": "Table 'Sales' not found in Data Model",
  "suggestedNextActions": ["table(action: 'add-to-data-model', table_name: 'Sales')"]
}
```

Always read and follow these suggestions before asking the user.

## Plugin-Specific Notes

- **Excel is opened by the plugin**: When the user selects a workbook via the control panel or `Excel VBA: 选择 Excel 文件`, the plugin automatically opens it through MCP. AI should not tell the user to manually open Excel.
- **Manual disconnection required**: The plugin does not automatically detect Excel closure. The user must click **断开 Excel** in the control panel to clean up the MCP session.
- **VBE sync is internal**: VBE ↔ local bidirectional sync is a plugin function handled by Legacy COM, not an MCP tool. AI should not use MCP `vba` tool actions for batch sync. See [vba-sync.md](vba-sync.md).
- **Window pinning is internal**: The plugin's `置顶` command uses legacy COM, not the MCP `window` tool. See [window.md](window.md) for MCP window operations.
- **Auto-sync is optional**: Enable in settings to automatically push local file changes to VBE. MCP-based writes do not trigger plugin auto-sync.

## Tool Selection Quick Reference

| Task | Tool | Key Action |
|------|------|------------|
| Create/open/save workbooks | `file` | open, create, close |
| Write/read cell data | `range` | set-values, get-values |
| Format cells (numbers) | `range` | set-number-format |
| Format cells (visual) | `range_format` | format-range, format-ranges |
| Create tables from data | `table` | create |
| Add table to Power Pivot | `table` | add-to-datamodel |
| Create DAX formulas | `datamodel` | create-measure |
| Create PivotTables | `pivottable` | create, create-from-datamodel |
| Filter with slicers | `slicer` | set-slicer-selection |
| Create charts | `chart` | create-from-range |
| Control calculation mode | `calculation_mode` | get-mode, set-mode, calculate |
| Visual verification | `screenshot` | capture, capture-sheet |
| Window visibility | `window` | show, arrange |

## Tool Categories

### 1. File Operations — `file`

| Action | Purpose |
|--------|---------|
| `open` | Open workbook, return `session_id`. |
| `close` | Close session with optional `save`. |
| `close-workbook` | Close workbook without quitting Excel. |
| `create-empty` | Create new `.xlsx` or `.xlsm`. |
| `list` | List active sessions. |
| `test` | Verify workbook is accessible; returns `is_irm_protected`. |

See [workbook-open.md](workbook-open.md) and [references/workflows.md](references/workflows.md).

### 2. Calculation Mode — `calculation_mode`

| Action | Purpose |
|--------|---------|
| `get-mode` | Query current mode/state. |
| `set-mode` | `automatic` / `manual` / `semi-automatic`. |
| `calculate` | Recalculate `workbook`, `sheet`, or `range`. |

See [calculation.md](calculation.md) and [references/calculation.md](references/calculation.md).

### 3. Worksheets — `worksheet`

| Action | Purpose |
|--------|---------|
| `list` | List sheets. |
| `create` | Create sheet (`before`/`after`). |
| `delete` | Delete sheet. |
| `rename` | Rename sheet. |
| `copy` / `move` | Copy/move within workbook or to `target_file`. |
| `set-tab-color` / `get-tab-color` / `clear-tab-color` | Tab colors. |
| `set-visibility` / `get-visibility` | `visible` / `hidden` / `very-hidden`. |

See [sheet.md](sheet.md) and [references/worksheet.md](references/worksheet.md).

### 4. Ranges — `range`

| Action | Purpose |
|--------|---------|
| `get-values` / `set-values` | Read/write cell values (2D arrays). |
| `get-formulas` / `set-formulas` | Read/write formulas. |
| `clear-contents` / `clear-all` / `clear-formats` | Clear range. |
| `get-used-range` / `get-current-region` / `get-range-info` | Discovery. |
| `copy` / `copy-values` / `copy-formulas` | Copy. |
| `insert-rows` / `delete-rows` / `insert-columns` / `delete-columns` / `insert-cells` / `delete-cells` | Structural edits. |
| `find` / `replace` | Search/replace. |
| `sort` | Sort. |
| `get-number-formats` / `set-number-format` / `set-number-formats` | Number formatting. |
| `add-hyperlink` / `remove-hyperlink` / `list-hyperlinks` / `get-hyperlink` | Hyperlinks. |

See [cell-read.md](cell-read.md), [cell-write.md](cell-write.md), and [references/range.md](references/range.md).

### 5. Range Format — `range_format`

| Action | Purpose |
|--------|---------|
| `get-style` / `set-style` | Built-in styles. |
| `set-format` / `format-ranges` | Font, fill, borders, alignment. |
| `add-validation` / `get-validation` / `remove-validation` | Data validation. |
| `merge-cells` / `unmerge-cells` / `get-merge-info` | Merge. |
| `set-cell-locked` / `get-cell-locked` | Protection. |
| `auto-fit-columns` / `auto-fit-rows` | Auto-fit. |

See [cell-write.md](cell-write.md) and [references/anti-patterns.md](references/anti-patterns.md).

### 6. Excel Tables — `table`

| Action | Purpose |
|--------|---------|
| `list` / `read` / `create` / `rename` / `resize` / `delete` | Lifecycle. |
| `apply-style` / `set-show-totals` / `set-column-totals` | Styling. |
| `get-data` / `append-rows` | Data ops. |
| `add-to-datamodel` | Required for DAX. |
| `create-from-dax` / `update-dax` / `get-dax` | DAX-backed tables. |
| `apply-filter` / `apply-filter-values` / `clear-filters` / `get-filter-state` | Filters. |
| `add-column` / `remove-column` / `rename-column` | Columns. |
| `get-structured-reference` | Formula syntax. |
| `sort` / `multi-column-sort` | Sorting. |
| `get-column-number-formats` / `set-column-number-formats` | Formatting. |

See [table.md](table.md) and [references/table.md](references/table.md).

### 7. PivotTables — `pivottable`

| Action | Purpose |
|--------|---------|
| `list` / `read` / `create` / `delete` | Lifecycle. |
| `create` (source_type `range`/`table`/`datamodel`) | Creation. |
| `add-field` / `remove-field` / `list-fields` | Field placement. |
| `set-field-function` / `set-field-name` / `set-field-number-format` / `set-field-filter` / `sort-field` | Field config. |
| `list-calculated-fields` / `create-calculated-field` / `delete-calculated-field` | Regular PivotTables. |
| `list-calculated-members` / `create-calculated-member` / `delete-calculated-member` | OLAP/datamodel PivotTables. |
| `set-layout` / `set-subtotals` / `set-grand-totals` | Layout. |
| `get-data` / `refresh` | Data ops. |

See [pivot-table.md](pivot-table.md) and [references/pivottable.md](references/pivottable.md).

### 8. Charts — `chart`

| Action | Purpose |
|--------|---------|
| `list` / `read` / `create` / `delete` / `move` | Lifecycle. |
| `create` (source_type `range`/`pivottable`) | Creation. |
| `add-series` / `remove-series` / `update-series-data` / `set-data-source` | Series. |
| `set-chart-type` / `set-legend-visible` / `set-style` | Configuration. |
| `set-title` / `set-axis-title` / `set-axis-number-format` / `get-axis-number-format` | Formatting. |
| `set-data-labels` / `set-label-position` | Data labels. |
| `get-axis-scale` / `set-axis-scale` | Axis scale. |
| `get-gridlines` / `set-gridlines` | Gridlines. |
| `set-marker-style` / `set-marker-size` / `set-marker-colors` | Series formatting. |
| `add-trendline` / `list-trendlines` / `delete-trendline` / `configure-trendline` | Trendlines. |
| `set-placement` / `fit-to-range` | Positioning. |

See [chart.md](chart.md), [references/chart.md](references/chart.md), and [references/dashboard.md](references/dashboard.md).

### 9. Power Query — `powerquery`

| Action | Purpose |
|--------|---------|
| `list` / `view` | Discovery. |
| `create` / `update` / `rename` / `delete` | Lifecycle. |
| `refresh` / `refresh-all` | Refresh. |
| `load-to` / `get-load-config` / `unload` | Load configuration. |
| `evaluate` | Test M code without persisting. |

See [power-query.md](power-query.md), [references/powerquery.md](references/powerquery.md), and [references/m-code-syntax.md](references/m-code-syntax.md).

### 10. Data Model / DAX — `datamodel`

| Action | Purpose |
|--------|---------|
| `list-tables` / `read-table` / `rename-table` / `delete-table` | Tables. |
| `list-columns` | Columns. |
| `list-measures` / `create-measure` / `update-measure` / `delete-measure` | Measures. |
| `read-info` | Model metadata. |
| `list-relationships` / `read-relationship` / `create-relationship` / `update-relationship` / `delete-relationship` | Relationships. |
| `refresh` | Refresh Data Model. |
| `list-workbook-connections` | Connections usable by Data Model. |
| `evaluate` | Run DAX `EVALUATE`. |
| `execute-dmv` | DMV metadata query. |

See [dax.md](dax.md), [references/datamodel.md](references/datamodel.md), and [references/dmv-reference.md](references/dmv-reference.md).

### 11. Named Ranges — `namedrange`

| Action | Purpose |
|--------|---------|
| `list` / `read` / `write` / `create` / `update` / `delete` | Full lifecycle. |

See [named-range.md](named-range.md).

### 12. Data Connections — `connection`

| Action | Purpose |
|--------|---------|
| `list` / `view` / `create` / `test` / `refresh` / `delete` | Lifecycle. |
| `load-to` | Load to worksheet. |
| `get-properties` / `set-properties` | Connection string & settings. |

See [connection.md](connection.md).

### 13. VBA — `vba`

| Action | Purpose |
|--------|---------|
| `list` / `view` | Read components and procedures. |
| `import` / `update` / `delete` | Modify components. |
| `run` | Execute procedure. |

See [vba-read.md](vba-read.md), [vba-write.md](vba-write.md), [vba-run.md](vba-run.md), and [vba-sync.md](vba-sync.md).

### 14. Slicers — `slicer`

| Action | Purpose |
|--------|---------|
| `list` / `create` / `set-selection` / `delete` | PivotTable slicers. |
| `list-table-slicers` / `create-table-slicer` / `set-table-selection` / `delete-table-slicer` | Table slicers. |

See [slicer.md](slicer.md) and [references/slicer.md](references/slicer.md).

### 15. Conditional Formatting — `conditionalformat`

| Action | Purpose |
|--------|---------|
| `add-rule` | Cell value, expression, color scale, data bar, icon set rules. |
| `clear-rules` | Remove rules from range. |

See [conditional-formatting.md](conditional-formatting.md) and [references/conditionalformat.md](references/conditionalformat.md).

### 16. Screenshot — `screenshot`

| Action | Purpose |
|--------|---------|
| `capture` | Capture range as PNG. |
| `capture-sheet` | Capture worksheet as PNG. |

See [screenshot.md](screenshot.md) and [references/screenshot.md](references/screenshot.md).

### 17. Window Management — `window`

| Action | Purpose |
|--------|---------|
| `show` / `hide` / `bring-to-front` | Visibility. |
| `get-info` | Window state/position. |
| `set-state` | `normal` / `minimized` / `maximized`. |
| `set-position` | Resize/move. |
| `arrange` | Presets: `left-half`, `right-half`, `top-half`, `bottom-half`, `center`, `full-screen`. |
| `set-status-bar` / `clear-status-bar` | Status bar text. |

See [window.md](window.md), [references/window.md](references/window.md), and [references/excel_agent_mode.md](references/excel_agent_mode.md).

## Important Rules

- **`session_id` is required** for all tools except `file(open)` and `file(list)`. Always use `file(open, path)` first to get a `session_id`.
- **Trust access**: Excel must have "Trust access to the VBA project object model" enabled, or all VBA read/write operations fail.
- **Close all other Excel instances** before automated operations to avoid file locking.
- **Use calculation mode for bulk writes**: when writing 10+ cells, set `manual`, write, calculate once, then restore `automatic`.
- **Prefer Excel Tables**: convert tabular data to tables for structured references, auto-expand, and Data Model/DAX support.
- **Data Model prerequisites**: add tables to Data Model with `table(add-to-datamodel)` before creating DAX measures.
- **Power Query best practice**: test M code with `powerquery(evaluate)` before creating persistent queries.
- **Document modules**: never try to create or delete worksheet/workbook object modules (`Sheet1`, `ThisWorkbook`). Only update their code via `vba(update)`.
- **UserForms**: do not create UserForms via `vba(import)` if you need the visual layout. Create them in VBE first, then update their code.
- **Follow `suggestedNextActions`** in error responses instead of asking the user.

## Dangerous Operations

These modify the workbook or execute code. Always verify before executing:

- `vba(import/update/delete/run)`
- `worksheet(create/delete/rename/copy/move/set-visibility)`
- `range(set-values/set-formulas/clear-*/insert/delete-rows)`
- `range_format(set-format/format-ranges/merge/unmerge/add-validation)`
- `table(create/delete/add-to-datamodel/resize/rename)`
- `pivottable(create/delete/add-field/remove-field/*)`
- `chart(create/delete/move)`
- `powerquery(create/update/delete)`
- `datamodel(create-measure/update-measure/delete-measure/create-relationship/*)`
- `namedrange(write/create/update/delete)`
- `connection(create/delete/refresh)`
- `slicer(create/delete/set-selection)`
- `conditionalformat(add-rule/clear-rules)`
- `window(set-position/set-state/show/hide)`
- Plugin sync commands: **VBE → 本地** and **本地 → VBE**

## Reference Documentation

See `references/` for detailed guidance:

- [Core execution rules and LLM guidelines](references/behavioral-rules.md)
- [Common mistakes to avoid](references/anti-patterns.md)
- [Bulk write performance optimization](references/calculation.md)
- [Data Model constraints and patterns](references/workflows.md)
- [Charts and formatting](references/chart.md)
- [Conditional formatting operations](references/conditionalformat.md)
- [Dashboard and report best practices](references/dashboard.md)
- [Data Model/DAX specifics](references/datamodel.md)
- [DMV query reference for Data Model analysis](references/dmv-reference.md)
- [Excel agent mode and advanced automation](references/excel_agent_mode.md)
- [Gotchas and known limits](references/gotchas.md)
- [Power Query M code syntax reference](references/m-code-syntax.md)
- [PivotTable operations](references/pivottable.md)
- [Power Query specifics](references/powerquery.md)
- [Range operations and number formats](references/range.md)
- [Screenshot and visual verification](references/screenshot.md)
- [Slicer operations](references/slicer.md)
- [Table operations](references/table.md)
- [Window and visibility operations](references/window.md)
- [Worksheet operations](references/worksheet.md)
