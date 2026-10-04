### conditionalformat

Conditional formatting - visual rules based on cell values

**Actions:** `update-rule`, `delete-rule`, `set-rule-priority`, `add-rule`, `clear-rules`, `list-rules`, `list-worksheet-rules`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet name, empty for active sheet (required) |
| `--rule-priority <RULEPRIORITY>` | Current worksheet-wide rule priority from listing (required for: update-rule, delete-rule, set-rule-priority) (valid for: update-rule, delete-rule, set-rule-priority) |
| `--expected-fingerprint <EXPECTEDFINGERPRINT>` | Fingerprint of the selected listed rule, not a persistent ID (required for: update-rule, delete-rule, set-rule-priority) (valid for: update-rule, delete-rule, set-rule-priority) |
| `--options <OPTIONS>` | Only supplied settings change; type-specific settings must match the existing rule type (required for: update-rule) (valid for: update-rule) (JSON format) |
| `--new-priority <NEWPRIORITY>` | New worksheet-wide priority from 1 through the greatest listed native priority (required for: set-rule-priority) (valid for: set-rule-priority) |
| `--range <RANGEADDRESS>` | Range address (A1 notation or named range) (required for: add-rule, clear-rules, list-rules) (valid for: add-rule, clear-rules, list-rules) |
| `--rule-type <RULETYPE>` | Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted. (required for: add-rule) (valid for: add-rule) |
| `--operator-type <OPERATORTYPE>` | Required for cellValue rules. XlFormatConditionOpe rator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween (valid for: add-rule) |
| `--formula1 <FORMULA1>` | Required for cellValue and expression rules. First formula/value for condition (valid for: add-rule) |
| `--formula2 <FORMULA2>` | Required for between/notBetween cellValue rules. Second formula/value (valid for: add-rule) |
| `--interior-color <INTERIORCOLOR>` | Fill color (#RRGGBB or color index) (valid for: add-rule) |
| `--interior-pattern <INTERIORPATTERN>` | Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.) (valid for: add-rule) |
| `--font-color <FONTCOLOR>` | Font color (#RRGGBB or color index) (valid for: add-rule) |
| `--font-bold <FONTBOLD>` | Bold font (valid for: add-rule) |
| `--font-italic <FONTITALIC>` | Italic font (valid for: add-rule) |
| `--border-style <BORDERSTYLE>` | Border style: none, continuous, dash, dot, etc. (valid for: add-rule) |
| `--border-color <BORDERCOLOR>` | Border color (#RRGGBB or color index) (valid for: add-rule) |
| `--color-scale-min-type <COLORSCALEMINTYPE>` | colorScale minimum stop type: minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-min-value <COLORSCALEMINVALUE>` | colorScale minimum stop value (for number/percent/perce ntile/formula) (valid for: add-rule) |
| `--color-scale-min-color <COLORSCALEMINCOLOR>` | colorScale minimum stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-mid-type <COLORSCALEMIDTYPE>` | colorScale midpoint stop type (supply to create a 3-color scale) (valid for: add-rule) |
| `--color-scale-mid-value <COLORSCALEMIDVALUE>` | colorScale midpoint stop value (valid for: add-rule) |
| `--color-scale-mid-color <COLORSCALEMIDCOLOR>` | colorScale midpoint stop color (#RRGGBB) (valid for: add-rule) |
| `--color-scale-max-type <COLORSCALEMAXTYPE>` | colorScale maximum stop type: maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--color-scale-max-value <COLORSCALEMAXVALUE>` | colorScale maximum stop value (valid for: add-rule) |
| `--color-scale-max-color <COLORSCALEMAXCOLOR>` | colorScale maximum stop color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-color <DATABARCOLOR>` | dataBar fill color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-negative-color <DATABARNEGATIVECOLOR>` | dataBar negative-value bar color (#RRGGBB) (valid for: add-rule) |
| `--data-bar-direction <DATABARDIRECTION>` | dataBar fill direction: context, leftToRight, rightToLeft (valid for: add-rule) |
| `--data-bar-show-value <DATABARSHOWVALUE>` | dataBar show the cell value alongside the bar (valid for: add-rule) |
| `--data-bar-min-type <DATABARMINTYPE>` | dataBar minimum point type: automaticMinimum, minimum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-min-value <DATABARMINVALUE>` | dataBar minimum point value (valid for: add-rule) |
| `--data-bar-max-type <DATABARMAXTYPE>` | dataBar maximum point type: automaticMaximum, maximum, number, percent, percentile, formula (valid for: add-rule) |
| `--data-bar-max-value <DATABARMAXVALUE>` | dataBar maximum point value (valid for: add-rule) |
| `--icon-set-id <ICONSETID>` | iconSet id: 3Arrows, 3TrafficLights1, 4Ratings, 5Quarters, etc. (valid for: add-rule) |
| `--icon-set-reverse <ICONSETREVERSE>` | iconSet reverse icon order (valid for: add-rule) |
| `--icon-set-show-icon-only <ICONSETSHOWICONONLY>` | iconSet show only the icon (hide the value) (valid for: add-rule) |
| `--icon-threshold1-type <ICONTHRESHOLD1TYPE>` | iconSet threshold 1 type: percent, number, percentile, formula (valid for: add-rule) |
| `--icon-threshold1-value <ICONTHRESHOLD1VALUE>` | iconSet threshold 1 value (valid for: add-rule) |
| `--icon-threshold2-type <ICONTHRESHOLD2TYPE>` | iconSet threshold 2 type (valid for: add-rule) |
| `--icon-threshold2-value <ICONTHRESHOLD2VALUE>` | iconSet threshold 2 value (valid for: add-rule) |
| `--icon-threshold3-type <ICONTHRESHOLD3TYPE>` | iconSet threshold 3 type (valid for: add-rule) |
| `--icon-threshold3-value <ICONTHRESHOLD3VALUE>` | iconSet threshold 3 value (valid for: add-rule) |
| `--icon-threshold4-type <ICONTHRESHOLD4TYPE>` | iconSet threshold 4 type (valid for: add-rule) |
| `--icon-threshold4-value <ICONTHRESHOLD4VALUE>` | iconSet threshold 4 value (valid for: add-rule) |
| `--rank <RANK>` | top10 rank (number of values, or percent when top10Percent is true) (valid for: add-rule) |
| `--top10-percent <TOP10PERCENT>` | top10 treat rank as a percentage (valid for: add-rule) |
| `--top-bottom <TOPBOTTOM>` | top10 direction: top or bottom (valid for: add-rule) |
| `--above-below <ABOVEBELOW>` | aboveAverage selector: aboveAverage, belowAverage, aboveStdDev, belowStdDev, equalAboveAverage, equalBelowAverage (valid for: add-rule) |
| `--date-period <DATEPERIOD>` | timePeriod period: today, yesterday, tomorrow, last7Days, thisWeek, lastWeek, nextWeek, thisMonth, lastMonth, nextMonth (valid for: add-rule) |
| `--priority <PRIORITY>` | Optional exact worksheet-wide priority for the new rule (valid for: add-rule) |
| `--stop-if-true <STOPIFTRUE>` | Optional evaluation stop flag; unavailable for colorScale/dataBar/i conSet rules (valid for: add-rule) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
