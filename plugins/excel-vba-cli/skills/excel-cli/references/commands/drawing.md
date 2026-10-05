### drawing

Worksheet drawing objects and sparklines

**Actions:** `group-objects`, `ungroup-object`, `align-objects`, `distribute-objects`, `duplicate-object`, `set-z-order`, `list-objects`, `get-object`, `add-image`, `add-shape`, `add-text-box`, `add-connector`, `add-form-control`, `update-object`, `delete-object`, `list-sparklines`, `get-sparkline`, `add-sparkline`, `update-sparkline`, `delete-sparkline`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet containing every selected object (required) |
| `--object-names <OBJECTNAMES>` | JSON array string of distinct top-level names returned by list-objects; at least two (required for: group-objects, align-objects, distribute-objects) (valid for: group-objects, align-objects, distribute-objects) |
| `--group-name <GROUPNAME>` | Optional unique name for the new group; otherwise keep Excel's generated name (valid for: group-objects) |
| `--object-name <OBJECTNAME>` | Existing top-level group name (required for: ungroup-object, duplicate-object, set-z-order, get-object, update-object, delete-object) (valid for: ungroup-object, duplicate-object, set-z-order, get-object, update-object, delete-object) |
| `--alignment <ALIGNMENT>` | Edges or centers to align: Left, Center, Right, Top, Middle, Bottom Accepted values (case-insensitive): Left, Center, Right, Top, Middle, Bottom. (required for: align-objects) (valid for: align-objects) |
| `--distribution <DISTRIBUTION>` | Horizontal or vertical spacing direction Accepted values (case-insensitive): Horizontal, Vertical. (required for: distribute-objects) (valid for: distribute-objects) |
| `--new-name <NEWNAME>` | Optional unique name; otherwise keep Excel's generated name (valid for: duplicate-object, update-object) |
| `--offset-left <OFFSETLEFT>` | Horizontal offset in points from the source, default 10 (valid for: duplicate-object) |
| `--offset-top <OFFSETTOP>` | Vertical offset in points from the source, default 10 (valid for: duplicate-object) |
| `--z-order <ZORDER>` | Native order change: BringToFront, SendToBack, BringForward, SendBackward Accepted values (case-insensitive): BringToFront, SendToBack, BringForward, SendBackward. (required for: set-z-order) (valid for: set-z-order) |
| `--image-path <IMAGEPATH>` | Full path to a readable local image file (required for: add-image) (valid for: add-image) |
| `--name <NAME>` | Optional name for the new drawing object (valid for: add-image, add-shape, add-text-box, add-connector, add-form-control) |
| `--left <LEFT>` | Left position in points from the worksheet edge (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--top <TOP>` | Top position in points from the worksheet edge (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--width <WIDTH>` | Object width in points (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--height <HEIGHT>` | Object height in points (valid for: add-image, add-shape, add-text-box, add-form-control, update-object) |
| `--lock-aspect-ratio <LOCKASPECTRATIO>` | Keep the image's aspect ratio when resizing (valid for: add-image) |
| `--shape-type <SHAPETYPE>` | AutoShape type, such as Rectangle, Oval, or a supported arrow/flowchart shape Accepted values (case-insensitive): Rectangle, Parallelogram, Trapezoid, Diamond, RoundedRectangle, Octagon, IsoscelesTriangle, RightTriangle, Oval, Hexagon, Cross, RegularPentagon, Can, Cube, Bevel, FoldedCorner, SmileyFace, Donut, NoSmoking, BlockArc, Heart, LightningBolt, Sun, Moon, Arc, RightArrow, LeftArrow, UpArrow, DownArrow, LeftRightArrow, UpDownArrow, QuadArrow, FlowchartProcess, FlowchartDecision, FlowchartData, FlowchartPredefinedProcess, FlowchartInternalStorage, FlowchartDocument, FlowchartMultidocument, FlowchartTerminator, FlowchartPreparation, FlowchartManualInput, FlowchartManualOperation, FlowchartConnector, FlowchartOffpageConnector. (valid for: add-shape) |
| `--text <TEXT>` | Text displayed by the shape, text box, or Forms control (required for: add-text-box) (valid for: add-shape, add-text-box, add-form-control, update-object) |
| `--fill-color <FILLCOLOR>` | Fill color as #RRGGBB (valid for: add-shape, add-text-box, update-object) |
| `--line-color <LINECOLOR>` | Outline, connector, or sparkline color as #RRGGBB (valid for: add-shape, add-text-box, add-connector, update-object, add-sparkline, update-sparkline) |
| `--line-weight <LINEWEIGHT>` | Line thickness in points (valid for: add-shape, add-connector, update-object) |
| `--font-size <FONTSIZE>` | Text size in points (valid for: add-text-box, update-object) |
| `--font-color <FONTCOLOR>` | Text color as #RRGGBB (valid for: add-text-box, update-object) |
| `--connector-type <CONNECTORTYPE>` | Connector geometry: Straight, Elbow, or Curved Accepted values (case-insensitive): Straight, Elbow, Curved. (valid for: add-connector) |
| `--begin-x <BEGINX>` | Starting horizontal position in points (valid for: add-connector) |
| `--begin-y <BEGINY>` | Starting vertical position in points (valid for: add-connector) |
| `--end-x <ENDX>` | Ending horizontal position in points (valid for: add-connector) |
| `--end-y <ENDY>` | Ending vertical position in points (valid for: add-connector) |
| `--control-type <CONTROLTYPE>` | Worksheet Forms control type; ActiveX/OLE controls are excluded Accepted values (case-insensitive): Button, CheckBox, DropDown, GroupBox, Label, ListBox, OptionButton, ScrollBar, Spinner. (valid for: add-form-control) |
| `--linked-cell <LINKEDCELL>` | Cell binding for CheckBox, DropDown, ListBox, OptionButton, ScrollBar, or Spinner (valid for: add-form-control, update-object) |
| `--input-range <INPUTRANGE>` | Cell range supplying items to a DropDown or ListBox (valid for: add-form-control, update-object) |
| `--rotation <ROTATION>` | Rotation angle in degrees (valid for: update-object) |
| `--visible <VISIBLE>` | Show or hide the object; omit to leave unchanged (valid for: update-object) |
| `--locked <LOCKED>` | Lock the object; effective when worksheet protection is enabled (valid for: update-object) |
| `--placement <PLACEMENT>` | Cell anchoring: 1=move and size, 2=move only, 3=free floating (valid for: update-object) |
| `--alternative-text <ALTERNATIVETEXT>` | Accessible description of the object (valid for: update-object) |
| `--location-range <LOCATIONRANGE>` | Cell or range displaying the sparkline group (required for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) (valid for: get-sparkline, add-sparkline, update-sparkline, delete-sparkline) |
| `--source-range <SOURCERANGE>` | Cell range supplying the sparkline data (required for: add-sparkline) (valid for: add-sparkline, update-sparkline) |
| `--sparkline-type <SPARKLINETYPE>` | Sparkline type: Line, Column, or WinLoss Accepted values (case-insensitive): Line, Column, WinLoss. (valid for: add-sparkline, update-sparkline) |
| `--show-markers <SHOWMARKERS>` | Display markers on line sparklines (valid for: add-sparkline, update-sparkline) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
