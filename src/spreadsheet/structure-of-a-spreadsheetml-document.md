# Structure of a SpreadsheetML document

A SpreadsheetML file is an Open Packaging Convention package. The `.xlsx` file is a ZIP container whose parts are connected by relationship items. `ooxmlsdk` exposes that graph through `SpreadsheetDocument` and generated part accessors.

## Package parts

| Package part | Root element | `ooxmlsdk` access |
|---|---|---|
| Workbook | `<workbook/>` | `SpreadsheetDocument::workbook_part()` |
| Worksheet | `<worksheet/>` | `WorkbookPart::worksheet_parts(&document)` |
| Chartsheet | `<chartsheet/>` | chartsheet parts when present |
| Shared strings | `<sst/>` | `WorkbookPart::shared_string_table_part(&document)` |
| Styles | `<styleSheet/>` | `WorkbookPart::workbook_styles_part(&document)` |
| Calculation chain | `<calcChain/>` | `WorkbookPart::calculation_chain_part(&document)` |
| Table | `<table/>` | `WorksheetPart::table_definition_parts(&document)` |
| Drawing | `<wsDr/>` | `WorksheetPart::drawings_part(&document)` |
| Pivot table | `<pivotTableDefinition/>` | `WorksheetPart::pivot_table_parts(&document)` |
| Pivot cache | `<pivotCacheDefinition/>` | workbook-level pivot cache parts when present |
| Pivot cache records | `<pivotCacheRecords/>` | pivot cache record parts when present |
| Conditional formatting | `<conditionalFormatting/>` | worksheet XML under the generated worksheet schema |

The exact set of parts depends on the workbook. A small workbook can contain only package relationships, `xl/workbook.xml`, one worksheet, and content type declarations.

The minimum workbook scenario has three spreadsheet-specific requirements: a single sheet entry, a workbook-local sheet ID, and a relationship ID that points from the workbook part to the worksheet part.

## Workbook and worksheet references

The workbook contains a `<sheets/>` collection. Each sheet entry has a workbook-local `sheetId` and a relationship id:

```xml
<workbook
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
</workbook>
```

The workbook relationship item resolves those ids to worksheet parts:

```xml
<Relationship
  Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  Target="worksheets/sheet1.xml"/>
```

## Reading the graph in Rust

Open the package, get the workbook part, and traverse typed child parts:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

That pattern is the base for the other SpreadsheetML examples. Prefer generated accessors over hard-coded ZIP paths when navigating relationships.

## Minimal worksheet

A worksheet stores rows and cells under `<sheetData/>`.

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>100</v></c>
    </row>
  </sheetData>
</worksheet>
```

Cells can store raw values, formulas, inline strings, or shared string indexes. The shared string table is a separate workbook-level part.

A typical workbook can add many more parts, including chartsheets, drawings, tables, pivot tables, pivot caches, styles, and shared strings. Keep the workbook XML, relationships, and content type declarations in sync whenever adding or removing parts.
