# slicer - Server Quirks

**Slicer Types**:

Two distinct slicer types exist:

- **PivotTable Slicers**: Filter PivotTables (can control multiple PivotTables)
- **Table Slicers**: Filter Excel Tables (single table only)

**Actions**:

| Action | Description | Required Parameters |
|--------|-------------|---------------------|
| `create-slicer` | Create PivotTable slicer | `pivot_table_name`, `field_name` |
| `list-slicers` | List all PivotTable slicers | (none) |
| `set-slicer-selection` | Set PivotTable slicer filter | `slicer_name`, `selected_items` |
| `delete-slicer` | Delete PivotTable slicer | `slicer_name` |
| `create-table-slicer` | Create Table slicer | `table_name`, `column_name` |
| `list-table-slicers` | List all Table slicers | (none) |
| `set-table-slicer-selection` | Set Table slicer filter | `slicer_name`, `selected_items` |
| `delete-table-slicer` | Delete Table slicer | `slicer_name` |

**CRITICAL: Required Parameters** - The "Required Parameters" column above is strict. Missing any required parameter will cause an error. Pay special attention to `pivot_table_name` for PivotTable slicers and `slicer_name` for selection/deletion operations.

**Naming Convention**:

- If `slicer_name` not provided, auto-generates `{FieldName}Slicer` or `{ColumnName}Slicer`
- Slicer names must be unique within workbook
- Use `list-slicers` or `list-table-slicers` to check existing names

**Selection Behavior**:

- `selected_items` is a list of strings: `["Value1", "Value2"]`
- Empty list `[]` clears all filters (shows all items)
- Values must match exactly (case-sensitive)
- Invalid values are silently ignored

**Positioning**:

- `destination_sheet` specifies which worksheet hosts the slicer
- `position` is a cell address for top-left corner (e.g., `'E1'`, `'G5'`)
- The slicer's top-left corner aligns to the specified cell
- Default position if not specified: Excel chooses

**Common Mistakes**:

- Creating slicer for field not in PivotTable → Error
- Creating table slicer for column not in table → Error
- Setting selection with wrong case → Values ignored (filter shows nothing)
- Deleting slicer that doesn't exist → Error

**Best Practices**:

1. Call `list-slicers` before creating to avoid name conflicts
2. Use `list-slicers` to get exact slicer names for selection/deletion
3. Multi-PivotTable filtering: Create one slicer, connect to multiple PivotTables in Excel UI
