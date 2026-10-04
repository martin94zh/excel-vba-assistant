### workbook

Manage workbook metadata, document properties, Save As/copy operations, fixed-format exports, and external Excel links

**Actions:** `list-table-styles`, `get-table-style`, `create-table-style`, `update-table-style`, `delete-table-style`, `list-cell-styles`, `get-cell-style`, `create-cell-style`, `update-cell-style`, `delete-cell-style`, `get-info`, `get-theme`, `apply-theme`, `list-document-properties`, `get-document-property`, `set-document-property`, `delete-document-property`, `save-as`, `save-copy-as`, `export-fixed-format`, `list-external-links`, `update-external-link`, `break-external-link`, `set-protection`, `get-protection`, `set-view-options`, `get-view-options`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--style-name <STYLENAME>` | Native cell or table style name. Discover existing names with list-cell-s tyles or list-table- styles; create actions require a new name. (required for: get-table-s tyle, create-tabl e-style, update-tabl e-style, delete-tabl e-style, get-cell-st yle, create-cell -style, update-cell -style, delete-cell -style) (valid for: get-table-s tyle, create-tabl e-style, update-tabl e-style, delete-tabl e-style, get-cell-st yle, create-cell -style, update-cell -style, delete-cell -style) |
| `--source-style-name <SOURCESTYLENAME>` | Existing native table/Pivot /slicer/tim eline style to clone. (required for: create-tabl e-style) (valid for: create-tabl e-style) |
| `--table-style-options <TABLESTYLEOPTIONS>` | Typed object: availabilit y flags and elements, each with elementType (native xl name), clear, stripeSize, bold/italic /underline/ strikethrou gh/themeFon t, font/fill color/theme /tint, and borders. Only row/column stripes accept stripeSize. No font name/size, script, alignment, number format or diagonals. (required for: update-tabl e-style) (valid for: update-tabl e-style) (JSON format) |
| `--source-sheet-name <SOURCESHEETNAME>` | Visible source worksheet name. (required for: create-cell -style) (valid for: create-cell -style) |
| `--source-cell-address <SOURCECELLADDRESS>` | Exactly one source cell in the selected workbook. (required for: create-cell -style) (valid for: create-cell -style) |
| `--style-options <STYLEOPTIONS>` | Typed object: formatOptio ns (same nested keys as range_forma t format, except inside borders are range-only) , includeFont /includeNum ber/include Alignment/i ncludeBorde r/includePa tterns/incl udeProtecti on, locked, formulaHidd en. Omitted inclusion flags are preserved. (required for: update-cell -style) (valid for: update-cell -style) (JSON format) |
| `--theme-path <THEMEPATH>` | Absolute path of the existing .thmx file. No theme file is created or copied. (required for: apply-theme ) (valid for: apply-theme ) |
| `--include-built-in <INCLUDEBUILTIN>` | Include built-in document properties (valid for: list-docume nt-properti es) |
| `--include-custom <INCLUDECUSTOM>` | Include custom document properties (valid for: list-docume nt-properti es) |
| `--property-name <PROPERTYNAME>` | Document property name (required for: get-documen t-property, set-documen t-property, delete-docu ment-proper ty) (valid for: get-documen t-property, set-documen t-property, delete-docu ment-proper ty) |
| `--scope <SCOPE>` | Property collection: built-in or custom Accepted values (case-insen sitive): BuiltIn, Custom. (valid for: get-documen t-property, set-documen t-property) |
| `--value <VALUE>` | String value to store (required for: set-documen t-property) (valid for: set-documen t-property) |
| `--target-path <TARGETPATH>` | Absolute output path in an existing directory (required for: save-as, save-copy-a s, export-fixe d-format) (valid for: save-as, save-copy-a s, export-fixe d-format) |
| `--format <FORMAT>` | Output format: auto, xlsx, xlsm, xlsb, or xls Accepted values (case-insen sitive): Auto, Xlsx, Xlsm, Xlsb, Xls. (valid for: save-as) |
| `--overwrite <OVERWRITE>` | Whether an existing output file may be replaced (valid for: save-as, save-copy-a s, export-fixe d-format) |
| `--format-type <FORMATTYPE>` | Fixed-forma t output: Pdf or Xps Accepted values (case-insen sitive): Pdf, Xps. (valid for: export-fixe d-format) |
| `--quality <QUALITY>` | Export quality: Standard or Minimum Accepted values (case-insen sitive): Standard, Minimum. (valid for: export-fixe d-format) |
| `--include-document-properties <INCLUDEDOCUMENTPROPERTIES>` | Include document metadata in the exported file (valid for: export-fixe d-format) |
| `--ignore-print-areas <IGNOREPRINTAREAS>` | Export without restricting output to configured print areas (valid for: export-fixe d-format) |
| `--from-page <FROMPAGE>` | First page to export, 1-based; omit to start at the beginning (valid for: export-fixe d-format) |
| `--to-page <TOPAGE>` | Last page to export, inclusive; omit to export through the end (valid for: export-fixe d-format) |
| `--open-after-publish <OPENAFTERPUBLISH>` | Open the exported file in its associated viewer (valid for: export-fixe d-format) |
| `--link-source <LINKSOURCE>` | Exact source identifier returned by list-extern al-links (required for: update-exte rnal-link, break-exter nal-link) (valid for: update-exte rnal-link, break-exter nal-link) |
| `--is-protected <ISPROTECTED>` | True to protect workbook structure, false to unprotect it (required for: set-protect ion) (valid for: set-protect ion) |
| `--password <PASSWORD>` | Optional protection password; required to unprotect password-pr otected structure (valid for: set-protect ion) |
| `--display-gridlines <DISPLAYGRIDLINES>` | Show or hide gridlines; omit to leave unchanged (valid for: set-view-op tions) |
| `--display-headings <DISPLAYHEADINGS>` | Show or hide row/column headings; omit to leave unchanged (valid for: set-view-op tions) |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
