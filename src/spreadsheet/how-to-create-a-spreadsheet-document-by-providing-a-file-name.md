# Create a spreadsheet document by providing a file name

Creating an `.xlsx` from scratch requires package relationships, content types, a workbook part, a workbook relationship item, and at least one worksheet part.

## Minimal package pieces

A minimal workbook includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `xl/workbook.xml`,
- `xl/workbook.xml`,
- `xl/_rels/workbook.xml.rels`,
- at least one `xl/worksheets/sheetN.xml`.

The `listings/spreadsheet` fixture builds this structure so the documented readers run against a real `.xlsx` package.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

This chapter does not yet publish a from-scratch writer. A final writer should validate relationship ids, sheet ids, content type overrides, workbook XML, worksheet XML, and save behavior.
