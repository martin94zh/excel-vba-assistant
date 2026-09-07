# chart

Create and manage Excel charts from ranges or PivotTables.

## Tool

`chart`

## Lifecycle

### List charts

```json
{
  "action": "list",
  "session_id": "abc123"
}
```

### Read chart info

```json
{
  "action": "read",
  "session_id": "abc123",
  "chartName": "SalesChart"
}
```

### Create chart from range

```json
{
  "action": "create",
  "session_id": "abc123",
  "sourceType": "range",
  "sourceAddress": "Sheet1!A1:D10",
  "chartType": "ColumnClustered",
  "destinationSheet": "ChartSheet",
  "destinationRange": "A3",
  "chartName": "SalesChart"
}
```

- `sourceType`: `range` | `pivottable`
- `chartType`: `ColumnClustered`, `BarClustered`, `Line`, `Pie`, `Scatter`, etc.
- `destinationSheet` / `destinationRange`: chart position

### Move chart

```json
{
  "action": "move",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "destinationSheet": "Dashboard",
  "destinationRange": "A1"
}
```

### Delete chart

```json
{
  "action": "delete",
  "session_id": "abc123",
  "chartName": "SalesChart"
}
```

## Series Management

### Add series

```json
{
  "action": "add-series",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Profit",
  "values": "Sheet1!$E$2:$E$10"
}
```

### Remove series

```json
{
  "action": "remove-series",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Profit"
}
```

### Update series data

```json
{
  "action": "update-series-data",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "values": "Sheet1!$B$2:$B$20",
  "categories": "Sheet1!$A$2:$A$20"
}
```

### Set data source

```json
{
  "action": "set-data-source",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "sourceAddress": "Sheet1!A1:D20"
}
```

## Configuration

### Set chart type

```json
{
  "action": "set-chart-type",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "chartType": "Line"
}
```

### Show/hide legend

```json
{
  "action": "set-legend-visible",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "visible": true
}
```

### Set chart style

```json
{
  "action": "set-style",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "styleId": 2
}
```

## Formatting

### Set chart title

```json
{
  "action": "set-title",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "title": "Monthly Sales"
}
```

### Set axis title

```json
{
  "action": "set-axis-title",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryCategory",
  "title": "Month"
}
```

- `axis`: `primaryCategory`, `primaryValue`, `secondaryCategory`, `secondaryValue`

### Set axis number format

```json
{
  "action": "set-axis-number-format",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue",
  "number_format": "$#,##0.00"
}
```

### Get axis number format

```json
{
  "action": "get-axis-number-format",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue"
}
```

## Data Labels

### Configure data labels

```json
{
  "action": "set-data-labels",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "showValue": true,
  "showPercentage": false,
  "showCategoryName": false
}
```

### Set label position

```json
{
  "action": "set-label-position",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "position": "OutsideEnd"
}
```

- `position`: `Center`, `InsideEnd`, `OutsideEnd`, etc.

## Axis Scale

### Get axis scale

```json
{
  "action": "get-axis-scale",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue"
}
```

### Set axis scale

```json
{
  "action": "set-axis-scale",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue",
  "minimum": 0,
  "maximum": 100000,
  "majorUnit": 10000,
  "minorUnit": 1000
}
```

## Gridlines

### Get gridlines config

```json
{
  "action": "get-gridlines",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue"
}
```

### Set gridlines

```json
{
  "action": "set-gridlines",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "axis": "primaryValue",
  "majorGridlines": true,
  "minorGridlines": false
}
```

## Series Formatting

### Set marker style

```json
{
  "action": "set-marker-style",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "style": "Circle"
}
```

- `style`: `Circle`, `Square`, `Diamond`, `Triangle`, etc.

### Set marker size

```json
{
  "action": "set-marker-size",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "size": 7
}
```

### Set marker colors

```json
{
  "action": "set-marker-colors",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "foregroundColor": "#FF0000",
  "backgroundColor": "#FFFFFF"
}
```

## Trendlines

### Add trendline

```json
{
  "action": "add-trendline",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "type": "Linear"
}
```

- `type`: `Linear`, `Exponential`, `Logarithmic`, `Polynomial`, `Power`, `MovingAverage`

### List trendlines

```json
{
  "action": "list-trendlines",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue"
}
```

### Delete trendline

```json
{
  "action": "delete-trendline",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "trendlineName": "Linear (Revenue)"
}
```

### Configure trendline

```json
{
  "action": "configure-trendline",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "seriesName": "Revenue",
  "trendlineName": "Linear (Revenue)",
  "forward": 6,
  "backward": 0,
  "displayEquation": true,
  "displayRSquared": true
}
```

## Placement & Positioning

### Set chart placement

```json
{
  "action": "set-placement",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "moveWithCells": false,
  "sizeWithCells": false
}
```

### Fit chart to range

```json
{
  "action": "fit-to-range",
  "session_id": "abc123",
  "chartName": "SalesChart",
  "sheet_name": "Dashboard",
  "range_address": "A1:K20"
}
```

## Notes

- For complex charts, create a basic chart first, then update series, titles, and formatting.
- Charts from PivotTables update automatically when the PivotTable refreshes.
- Use `screenshot` to verify chart appearance after styling changes.
