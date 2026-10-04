### screenshot

Capture Excel worksheet content as images for visual verification

**Actions:** `capture`, `capture-sheet`

| Parameter | Description |
|-----------|-------------|
| `--session <SESSION>` | Session ID from 'session open' command |
| `--sheet <SHEETNAME>` | Worksheet name (null for active sheet) |
| `--range <RANGEADDRESS>` | Range to capture (e.g., "A1:F20") (valid for: capture) |
| `--quality <QUALITY>` | Image quality: Medium (default, JPEG 75% scale), High (PNG full scale), Low (JPEG 50% scale) Accepted values (case-insensitive): Medium, High, Low |
| `--output <PATH>` | Write output to file instead of stdout. For image results, decodes and saves as binary file |
