### chart

Chart lifecycle - create, read, move, and delete embedded charts

**Actions:** `export-image`, `list`, `read`, `create-from-range`, `create-from-table`, `create-from-pivottable`, `delete`, `move`, `fit-to-range`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--chart-name <CHARTNAME>` | Existing embedded chart name (required for: export-image, read, delete, move, fit-to-range) (valid for: export-image, read, create-from-range, create-from-table, create-from-pivottable, delete, move, fit-to-range) |
| `--target-path <TARGETPATH>` | Output image path in an existing directory; extension must match the requested image format (required for: export-image) (valid for: export-image) |
| `--image-format <IMAGEFORMAT>` | Png (default), Jpeg or Gif; available filters depend on installed Excel Accepted values (case-insensitive): Png, Jpeg, Gif. (valid for: export-image) |
| `--overwrite <OVERWRITE>` | Explicit permission to replace an existing output image; default false (valid for: export-image) |
| `--sheet <SHEETNAME>` | Target worksheet name (required for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) (valid for: create-from-range, create-from-table, create-from-pivottable, fit-to-range) |
| `--source-range-address <SOURCERANGEADDRESS>` | Data range for the chart (e.g., A1:D10) (required for: create-from-range) (valid for: create-from-range) |
| `--chart-type <CHARTTYPE>` | Type of chart to create Accepted values (case-insensitive): ColumnClustered, ColumnStacked, ColumnStacked100, Column3DClustered, Column3DStacked, Column3DStacked100, Column3D, BarClustered, BarStacked, BarStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, Line3D, Pie, Pie3D, PieOfPie, PieExploded, PieExploded3D, BarOfPie, XYScatter, XYScatterSmooth, XYScatterSmoothNoMarkers, XYScatterLines, XYScatterLinesNoMarkers, Area, AreaStacked, AreaStacked100, Area3D, Area3DStacked, Area3DStacked100, Doughnut, DoughnutExploded, Radar, RadarMarkers, RadarFilled, Surface, SurfaceWireframe, SurfaceTopView, SurfaceTopViewWireframe, Bubble, Bubble3DEffect, StockHLC, StockOHLC, StockVHLC, StockVOHLC, CylinderBarClustered, CylinderBarStacked, CylinderBarStacked100, CylinderCol, CylinderColClustered, CylinderColStacked, CylinderColStacked100, ConeBarClustered, ConeBarStacked, ConeBarStacked100, ConeCol, ConeColClustered, ConeColStacked, ConeColStacked100, PyramidBarClustered, PyramidBarStacked, PyramidBarStacked100, PyramidCol, PyramidColClustered, PyramidColStacked, PyramidColStacked100, Treemap, Sunburst, Histogram, Pareto, BoxWhisker, Waterfall, Funnel, ColumnLineCombo, RegionMap. (required for: create-from-range, create-from-table, create-from-pivottable) (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--left <LEFT>` | Left position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--top <TOP>` | Top position in points from worksheet edge (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--width <WIDTH>` | Chart width in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--height <HEIGHT>` | Chart height in points (valid for: create-from-range, create-from-table, create-from-pivottable, move) |
| `--target-range <TARGETRANGE>` | Cell range to position chart within (e.g., 'F2:K15'). PREFERRED over left/top. When set, left/top are ignored. (valid for: create-from-range, create-from-table, create-from-pivottable) |
| `--table-name <TABLENAME>` | Name of the Excel Table (required for: create-from-table) (valid for: create-from-table) |
| `--pivot-table-name <PIVOTTABLENAME>` | Name of the source PivotTable (required for: create-from-pivottable) (valid for: create-from-pivottable) |
| `--range <RANGEADDRESS>` | Range to fit the chart to (e.g., A1:D10) (required for: fit-to-range) (valid for: fit-to-range) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
