# CLI Command Reference

> Auto-generated from the built `excelcli` runtime. Use these exact command and parameter names.

## Common Pitfalls

- `--values-file` requires an existing JSON or CSV file; use `--values` for inline JSON.
- `--timeout` ranges are action-specific: session open/create accepts 10-3600; Power Query refresh/refresh-all accepts 0-2147483 (0 keeps the default); other generated timeout actions accept 1-2147483.
- `pythoninexcel get-result --max-wait-seconds` must be at least 1 and shorter than the session operation timeout.
- `--values` and list parameters use JSON arrays; range values use a two-dimensional array.
- Power Query operations may take 30 seconds or longer; use a deliberate data-operation timeout or 0 for the default.
