### slicer

Slicer visual filters for PivotTables and Excel Tables

**Actions:** `create-timeline`, `get-slicer`, `update-slicer`, `set-timeline-selection`, `clear-timeline-selection`, `connect-pivottable`, `disconnect-pivottable`, `create-slicer`, `list-slicers`, `set-slicer-selection`, `delete-slicer`, `create-table-slicer`, `list-table-slicers`, `set-table-slicer-selection`, `delete-table-slicer`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--pivot-table-name <PIVOTTABLENAME>` | Source PivotTable (required for: create-timeline, connect-pivottable, disconnect-pivottable, create-slicer) (valid for: create-timeline, connect-pivottable, disconnect-pivottable, create-slicer, list-slicers) |
| `--field-name <FIELDNAME>` | Native date field (required for: create-timeline, create-slicer) (valid for: create-timeline, create-slicer) |
| `--slicer-name <SLICERNAME>` | Unique timeline name (required for: create-timeline, get-slicer, update-slicer, set-timeline-selection, clear-timeline-selection, connect-pivottable, disconnect-pivottable, create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) (valid for: create-timeline, get-slicer, update-slicer, set-timeline-selection, clear-timeline-selection, connect-pivottable, disconnect-pivottable, create-slicer, set-slicer-selection, delete-slicer, create-table-slicer, set-table-slicer-selection, delete-table-slicer) |
| `--destination-sheet <DESTINATIONSHEET>` | Destination worksheet (required for: create-timeline, create-slicer, create-table-slicer) (valid for: create-timeline, create-slicer, create-table-slicer) |
| `--position <POSITION>` | Single top-left anchor cell (required for: create-timeline, create-slicer, create-table-slicer) (valid for: create-timeline, create-slicer, create-table-slicer) |
| `--slicer-options <SLICEROPTIONS>` | Typed native layout settings; nested JSON keys use camelCase (required for: update-slicer) (valid for: update-slicer) |
| `--timeline-selection <TIMELINESELECTION>` | Required startDate and endDate calendar dates, in ascending order (required for: set-timeline-selection) (valid for: set-timeline-selection) (JSON format) |
| `--selected-items <SELECTEDITEMS>` | Items to select (show in PivotTable); empty clears the filter. Data Model/OLAP accepts returned captions or MDX unique names; unknown or ambiguous items fail without changing the filter. (required for: set-slicer-selection, set-table-slicer-selection) (valid for: set-slicer-selection, set-table-slicer-selection) (JSON format) |
| `--clear-first <CLEARFIRST>` | If true, clears existing selection before setting new items (default: true) (valid for: set-slicer-selection, set-table-slicer-selection) |
| `--table-name <TABLENAME>` | Name of the Excel Table (required for: create-table-slicer) (valid for: create-table-slicer, list-table-slicers) |
| `--column-name <COLUMNNAME>` | Name of the column to use for the slicer (required for: create-table-slicer) (valid for: create-table-slicer) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
