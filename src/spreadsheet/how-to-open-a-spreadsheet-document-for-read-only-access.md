# Open a spreadsheet document for read-only access

Use `SpreadsheetDocument` to open an `.xlsx` package and inspect its workbook and worksheet parts. In `ooxmlsdk`, opening a package does not modify it; changes are persisted only when you call a save method.

For read-only inspection, open the package and avoid saving it.

## Open and inspect worksheet parts

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

The example uses lazy package opening:

- `OpenSettings { open_mode: PackageOpenMode::Lazy, ..Default::default() }`
- `SpreadsheetDocument::new_from_file_with_settings`
- `workbook_part()`
- `worksheet_parts(&document)`

Lazy opening is useful for inspection helpers because it lets you navigate the package model without parsing every root element up front.

## Spreadsheet package structure

A SpreadsheetML package stores the main workbook in `xl/workbook.xml`. Worksheets are separate parts, usually under `xl/worksheets/`, and the workbook part owns relationships to those worksheet parts.

Use relationships and typed part accessors instead of hard-coding ZIP paths whenever possible.
