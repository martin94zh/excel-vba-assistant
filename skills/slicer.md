# slicer

Create and control slicers for PivotTables and Excel Tables.

## Tool

`slicer`

## Parameters

### List slicers

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Create PivotTable slicer

```json
{
  "action": "create",
  "session_id": "abc123",
  "slicerType": "pivottable",
  "pivotTableName": "SalesPivot",
  "fieldName": "Region",
  "slicerName": "RegionSlicer",
  "destinationSheet": "Dashboard",
  "destinationRange": "A1"
}
```

### Create Table slicer

```json
{
  "action": "create",
  "session_id": "abc123",
  "slicerType": "table",
  "tableName": "SalesData",
  "columnName": "Region",
  "slicerName": "RegionSlicer",
  "destinationSheet": "Dashboard",
  "destinationRange": "A1"
}
```

### Set selection

```json
{
  "action": "set-selection",
  "session_id": "abc123",
  "slicerName": "RegionSlicer",
  "selection": ["North", "South"]
}
```

### Delete slicer

```json
{
  "action": "delete",
  "session_id": "abc123",
  "slicerName": "RegionSlicer"
}
```

## Notes

- Slicers provide interactive filtering without modifying underlying structure.
- Multi-select is supported by passing an array of values.
