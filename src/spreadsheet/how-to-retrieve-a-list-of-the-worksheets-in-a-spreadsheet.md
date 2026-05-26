# Retrieve a list of the worksheets in a spreadsheet document

Worksheet names are stored in the workbook part, not inside the worksheet parts. The workbook `<sheets/>` collection maps each visible workbook entry to a relationship id.

## Read worksheet names

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

When you need the relationship ids as well as the sheet names, use `related_parts_of_type` at the workbook part boundary and keep those ids aligned with the `r:id` values in workbook XML:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheet_relationship_ids}}
```

The helper opens the workbook part, reads `xl/workbook.xml`, and extracts each `<sheet name="..."/>` value. The same workbook XML also contains `sheetId`, `r:id`, and optional state attributes such as `hidden`.

The returned list can be empty only for an invalid or unusual workbook; a well-formed spreadsheet normally has at least one sheet entry.

## Workbook markup

```xml
<sheets>
  <sheet name="Summary" sheetId="1" r:id="rId1"/>
  <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
```

Use `WorkbookPart::worksheet_parts(&document)` when you need the actual worksheet XML. Use workbook XML when you need workbook metadata such as names and visibility state.

The workbook `sheet` element is metadata, not the worksheet part itself. Resolve the `r:id` relationship when you need to move from a listed sheet to its worksheet XML.
