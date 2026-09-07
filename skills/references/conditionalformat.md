# conditionalformat - Server Quirks

**Rule Types**:

| Type | Description | Parameters |
|------|-------------|------------|
| `cell-value` | Format based on cell value comparison | operator_type + formula1 (+ formula2 for between) |
| `expression` | Format based on formula result | formula only |

**Operators (for cell-value type)**:

| Operator | Description | Formulas Required |
|----------|-------------|-------------------|
| `equal` | Cell equals value | formula1 |
| `not-equal` | Cell doesn't equal value | formula1 |
| `greater` | Cell greater than value | formula1 |
| `less` | Cell less than value | formula1 |
| `greater-equal` | Cell greater or equal | formula1 |
| `less-equal` | Cell less or equal | formula1 |
| `between` | Cell between two values | formula1 AND formula2 |
| `not-between` | Cell not between two values | formula1 AND formula2 |

**Format Options**:

- `interior_color`: Background fill color as `#RRGGBB` hex
- `font_color`: Text color as `#RRGGBB` hex
- `font_bold`: `true` or `false`
- `font_italic`: `true` or `false`
- `border_style`: Excel border style name
- `border_color`: Border color as `#RRGGBB` hex

**Actions**:

| Action | Description |
|--------|-------------|
| `add-rule` | Add conditional formatting rule to range |
| `clear-rules` | Remove all conditional formatting from range |

**Formula Notes**:

- For `cell-value` type: formula1/formula2 can be numbers, strings, or cell references
- For `expression` type: formula must return TRUE/FALSE
- Formulas use the top-left cell perspective (e.g., `=$A1>100` for relative rows)
- Use absolute references (`$A$1`) when comparing to a fixed cell

**Examples**:

**Highlight cells greater than 100:**

```json
{
  "action": "add-rule",
  "range_address": "A1:A10",
  "rule_type": "cell-value",
  "operator_type": "greater",
  "formula1": "100",
  "interior_color": "#FFFF00"
}
```

**Highlight cells between 50 and 100:**

```json
{
  "action": "add-rule",
  "range_address": "A1:A10",
  "rule_type": "cell-value",
  "operator_type": "between",
  "formula1": "50",
  "formula2": "100",
  "interior_color": "#90EE90"
}
```

**Highlight row if column A is "Active" (expression):**

```json
{
  "action": "add-rule",
  "range_address": "A1:D10",
  "rule_type": "expression",
  "formula": "=$A1=\"Active\"",
  "interior_color": "#90EE90"
}
```

**Common Mistakes**:

- Using `cell-value` type without `operator_type` → Error
- Using `between` without both formula1 AND formula2 → Error
- Forgetting `$` in expression formulas → Rule applies incorrectly across rows/columns
- Colors without `#` prefix → May not apply correctly

**Best Practices**:

1. Test expression formulas in Excel first to verify logic
2. Use `clear-rules` before applying new rules if replacing existing formatting
3. For row-based highlighting, apply rule to full range (not just one column)
4. Use relative row references (`$A1`) and absolute column references for row highlighting
