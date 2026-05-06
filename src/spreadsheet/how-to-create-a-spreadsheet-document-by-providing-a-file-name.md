# Create a spreadsheet document by providing a file name

Creating an `.xlsx` from scratch requires package relationships, content types, a workbook part, a workbook relationship item, and at least one worksheet part.

In `ooxmlsdk 0.6.1`, create the package with `SpreadsheetDocument::create(SpreadsheetDocumentType::Workbook)`, add the workbook part, add worksheet parts from the workbook part, write the root XML, and save the package to the file or writer your application owns.

Choose the package type and file extension together. A normal workbook uses `.xlsx`; macro-enabled workbooks, templates, and add-ins use different extensions and content types. Excel can reject a file when the package type and extension do not match.

## Minimal package pieces

A minimal workbook includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `xl/workbook.xml`,
- `xl/workbook.xml`,
- `xl/_rels/workbook.xml.rels`,
- at least one `xl/worksheets/sheetN.xml`.

This minimal writer creates that structure in memory:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:create_spreadsheet_document}}
```

At the SpreadsheetML layer, the workbook root owns the `sheets` collection. Each `sheet` entry stores the display name, a workbook-local `sheetId`, and an `r:id` relationship to the worksheet part. The worksheet part itself owns the `worksheet` root and `sheetData`.

For template workflows, `SpreadsheetDocument::create_from_template` opens an `.xltx` or `.xltm` as an editable workbook package. Use `document_type()` to inspect the current type and `change_document_type(...)` when converting the package content type deliberately.
