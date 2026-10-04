### chartconfig

Chart configuration - data source, series, type, title, axis labels, legend, and styling

**Actions:** `get-error-bars`, `set-error-bars`, `get-point-format`, `set-point-format`, `get-series-settings`, `set-series-axis-group`, `set-source-range`, `add-series`, `remove-series`, `set-chart-type`, `set-title`, `set-axis-title`, `get-axis-number-format`, `set-axis-number-format`, `show-legend`, `set-style`, `set-placement`, `set-data-labels`, `get-axis-scale`, `set-axis-scale`, `get-gridlines`, `set-gridlines`, `set-series-format`, `set-series-chart-type`, `get-plot-options`, `set-plot-options`, `set-area-format`, `list-trendlines`, `add-trendline`, `delete-trendline`, `set-trendline`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--chart-name <CHARTNAME>` | Existing embedded chart name (required) |
| `--series-index <SERIESINDEX>` | One-based existing series index (required for: get-error-bars, set-error-bars, get-point-format, set-point-format, get-series-settings , set-series-axis-gro up, remove-series, set-series-format, set-series-chart-ty pe, list-trendlines, add-trendline, delete-trendline, set-trendline) (valid for: get-error-bars, set-error-bars, get-point-format, set-point-format, get-series-settings , set-series-axis-gro up, remove-series, set-data-labels, set-series-format, set-series-chart-ty pe, list-trendlines, add-trendline, delete-trendline, set-trendline) |
| `--error-bar-options <ERRORBAROPTIONS>` | Typed bar settings using camelCase nested keys; custom ranges are on sourceSheetName or the chart worksheet (required for: set-error-bars) (valid for: set-error-bars) (JSON format) |
| `--point-index <POINTINDEX>` | One-based existing point index (required for: get-point-format, set-point-format) (valid for: get-point-format, set-point-format) |
| `--point-options <POINTOPTIONS>` | Typed point material/marker settings with camelCase nested keys; omitted fields preserve existing settings (required for: set-point-format) (valid for: set-point-format) (JSON format) |
| `--axis-group <AXISGROUP>` | Primary or Secondary; native chart type must support the requested assignment Accepted values (case-insensitive): Primary, Secondary. (required for: set-series-axis-gro up) (valid for: set-series-axis-gro up) |
| `--source-range <SOURCERANGE>` | New data source range (e.g., Sheet1!A1:D10) (required for: set-source-range) (valid for: set-source-range) |
| `--series-name <SERIESNAME>` | Display name for the series (required for: add-series) (valid for: add-series) |
| `--values-range <VALUESRANGE>` | Series values on the chart worksheet (e.g., B2:B10), or a sheet-qualified range in the chart workbook (e.g., Sheet1!B2:B10) (required for: add-series) (valid for: add-series) |
| `--category-range <CATEGORYRANGE>` | Optional category labels on the chart worksheet (e.g., A2:A10), or a sheet-qualified range in the chart workbook (valid for: add-series) |
| `--chart-type <CHARTTYPE>` | New chart type to apply Accepted values (case-insensitive): ColumnClustered, ColumnStacked, ColumnStacked100, Column3DClustered, Column3DStacked, Column3DStacked100, Column3D, BarClustered, BarStacked, BarStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked1 00, Line3D, Pie, Pie3D, PieOfPie, PieExploded, PieExploded3D, BarOfPie, XYScatter, XYScatterSmooth, XYScatterSmoothNoMa rkers, XYScatterLines, XYScatterLinesNoMar kers, Area, AreaStacked, AreaStacked100, Area3D, Area3DStacked, Area3DStacked100, Doughnut, DoughnutExploded, Radar, RadarMarkers, RadarFilled, Surface, SurfaceWireframe, SurfaceTopView, SurfaceTopViewWiref rame, Bubble, Bubble3DEffect, StockHLC, StockOHLC, StockVHLC, StockVOHLC, CylinderBarClustere d, CylinderBarStacked, CylinderBarStacked1 00, CylinderCol, CylinderColClustere d, CylinderColStacked, CylinderColStacked1 00, ConeBarClustered, ConeBarStacked, ConeBarStacked100, ConeCol, ConeColClustered, ConeColStacked, ConeColStacked100, PyramidBarClustered , PyramidBarStacked, PyramidBarStacked10 0, PyramidCol, PyramidColClustered , PyramidColStacked, PyramidColStacked10 0, Treemap, Sunburst, Histogram, Pareto, BoxWhisker, Waterfall, Funnel, ColumnLineCombo, RegionMap. (required for: set-chart-type, set-series-chart-ty pe) (valid for: set-chart-type, set-series-chart-ty pe) |
| `--title <TITLE>` | Title text to display (required for: set-title, set-axis-title) (valid for: set-title, set-axis-title) |
| `--axis <AXIS>` | Axis selector: Category/Value for primary axes, CategorySecondary/V alueSecondary for secondary axes. Primary aliases Category; Secondary aliases Value on the primary group. The requested axis must exist. Accepted values (case-insensitive): Primary, Secondary, Category, Value, CategorySecondary, ValueSecondary. (required for: set-axis-title, get-axis-number-for mat, set-axis-number-for mat, get-axis-scale, set-axis-scale, set-gridlines) (valid for: set-axis-title, get-axis-number-for mat, set-axis-number-for mat, get-axis-scale, set-axis-scale, set-gridlines) |
| `--number-format <NUMBERFORMAT>` | Excel number format code (e.g., "$#,##0", "0.00%") (required for: set-axis-number-for mat) (valid for: set-axis-number-for mat) |
| `--visible <VISIBLE>` | True to show legend, false to hide (required for: show-legend) (valid for: show-legend) |
| `--legend-position <LEGENDPOSITION>` | Optional position for the legend Accepted values (case-insensitive): Bottom, Corner, Custom, Left, Right, Top. (valid for: show-legend) |
| `--style-id <STYLEID>` | Excel chart style ID (1-48 for most chart types) (required for: set-style) (valid for: set-style) |
| `--placement <PLACEMENT>` | Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating (required for: set-placement) (valid for: set-placement) |
| `--print-object <PRINTOBJECT>` | Whether the embedded chart prints with the worksheet (valid for: set-placement) |
| `--locked <LOCKED>` | Whether the embedded chart is locked when the worksheet is protected (valid for: set-placement) |
| `--rounded-corners <ROUNDEDCORNERS>` | Whether the embedded chart uses rounded corners (valid for: set-placement) |
| `--show-value <SHOWVALUE>` | Show data values on labels (valid for: set-data-labels) |
| `--show-percentage <SHOWPERCENTAGE>` | Show percentage values for pie and doughnut charts. Excel can reject this setting on other chart types; rejection is an error, not a successful no-op. (valid for: set-data-labels) |
| `--show-series-name <SHOWSERIESNAME>` | Show series name on labels (valid for: set-data-labels) |
| `--show-category-name <SHOWCATEGORYNAME>` | Show category name on labels (valid for: set-data-labels) |
| `--show-bubble-size <SHOWBUBBLESIZE>` | Show bubble size (bubble charts) (valid for: set-data-labels) |
| `--separator <SEPARATOR>` | Separator string between label components (valid for: set-data-labels) |
| `--label-position <LABELPOSITION>` | Position of data labels relative to data points Accepted values (case-insensitive): BestFit, Center, Above, Below, Left, Right, InsideBase, InsideEnd, OutsideEnd. (valid for: set-data-labels) |
| `--minimum-scale <MINIMUMSCALE>` | Minimum axis value (null for auto) (valid for: set-axis-scale) |
| `--maximum-scale <MAXIMUMSCALE>` | Maximum axis value (null for auto) (valid for: set-axis-scale) |
| `--major-unit <MAJORUNIT>` | Major gridline interval (null for auto) (valid for: set-axis-scale) |
| `--minor-unit <MINORUNIT>` | Minor gridline interval (null for auto) (valid for: set-axis-scale) |
| `--show-major <SHOWMAJOR>` | Show major gridlines (null to keep current) (valid for: set-gridlines) |
| `--show-minor <SHOWMINOR>` | Show minor gridlines (null to keep current) (valid for: set-gridlines) |
| `--marker-style <MARKERSTYLE>` | Marker shape style Accepted values (case-insensitive): None, Automatic, Circle, Dash, Diamond, Dot, Picture, Plus, Square, Star, Triangle, X. (valid for: set-series-format) |
| `--marker-size <MARKERSIZE>` | Marker size in points (2-72) (valid for: set-series-format) |
| `--marker-background-color <MARKERBACKGROUNDCOLOR>` | Marker fill color (#RRGGBB) (valid for: set-series-format) |
| `--marker-foreground-color <MARKERFOREGROUNDCOLOR>` | Marker border color (#RRGGBB) (valid for: set-series-format) |
| `--invert-if-negative <INVERTIFNEGATIVE>` | Invert colors for negative values (valid for: set-series-format) |
| `--fill-color <FILLCOLOR>` | Series fill color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--fill-transparency <FILLTRANSPARENCY>` | Series fill transparency from 0 (opaque) to 1 (transparent) (valid for: set-series-format, set-area-format) |
| `--line-color <LINECOLOR>` | Series line color as #RRGGBB (valid for: set-series-format, set-area-format) |
| `--line-weight <LINEWEIGHT>` | Series line weight in points (valid for: set-series-format, set-area-format) |
| `--plot-by <PLOTBY>` | Interpret source rows or columns as data series Accepted values (case-insensitive): Rows, Columns. (valid for: set-plot-options) |
| `--display-blanks-as <DISPLAYBLANKSAS>` | How blank cells appear: gaps, zeroes, or interpolation Accepted values (case-insensitive): Gaps, Zero, Interpolated. (valid for: set-plot-options) |
| `--plot-visible-only <PLOTVISIBLEONLY>` | True to omit hidden rows and columns (valid for: set-plot-options) |
| `--area <AREA>` | Chart area or plot area Accepted values (case-insensitive): Chart, Plot. (required for: set-area-format) (valid for: set-area-format) |
| `--trendline-type <TRENDLINETYPE>` | Type of trendline (Linear, Exponential, etc.) Accepted values (case-insensitive): Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage. (required for: add-trendline) (valid for: add-trendline) |
| `--order <ORDER>` | Polynomial order (2-6, for Polynomial type) (valid for: add-trendline) |
| `--period <PERIOD>` | Moving average period (for MovingAverage type) (valid for: add-trendline) |
| `--forward <FORWARD>` | Periods to extend forward (valid for: add-trendline, set-trendline) |
| `--backward <BACKWARD>` | Periods to extend backward (valid for: add-trendline, set-trendline) |
| `--intercept <INTERCEPT>` | Force trendline through specific Y-intercept (valid for: add-trendline, set-trendline) |
| `--display-equation <DISPLAYEQUATION>` | Display trendline equation on chart (valid for: add-trendline, set-trendline) |
| `--display-r-squared <DISPLAYRSQUARED>` | Display R-squared value on chart (valid for: add-trendline, set-trendline) |
| `--name <NAME>` | Custom name for the trendline (valid for: add-trendline, set-trendline) |
| `--trendline-index <TRENDLINEINDEX>` | 1-based index of the trendline to delete (required for: delete-trendline, set-trendline) (valid for: delete-trendline, set-trendline) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
