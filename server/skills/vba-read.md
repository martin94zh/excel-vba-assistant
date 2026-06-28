# vba-read

Read VBAProject structure and code without modifying anything.

## Tools

- `excel_list_resources`
- `excel_get_vba_code`
- `excel_get_all_vba_code`
- `excel_list_macros`

## Parameters

### excel_list_resources

```json
{ "workbookId": "book" }
```

Returns every VBComponent with `name`, `type` (`StdModule` / `ClassModule` / `UserForm` / `Worksheet` / `ThisWorkbook`), and `codeLines`.

### excel_get_vba_code

```json
{
  "workbookId": "book",
  "componentName": "Module1"
}
```

- `componentName` (required): exact VBE component name.

### excel_get_all_vba_code

```json
{ "workbookId": "book" }
```

Returns the full code of every component in one call.

### excel_list_macros

```json
{ "workbookId": "book" }
```

Returns Sub/Function names in `module.procedure` form.

## Notes

- Use `excel_get_all_vba_code` instead of looping through `excel_get_vba_code` when you need to understand the whole project.
- Use `excel_list_macros` to verify a macro exists before running it.
