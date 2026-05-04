# Working with sheets

Sheets are workbook entries that point to sheet parts. The most common sheet type is a worksheet, which stores a grid of rows and cells. Workbooks can also contain chartsheets, dialog sheets, and macro sheets.

## Workbook sheet list

The workbook part stores sheet metadata:

```xml
<sheets>
  <sheet name="Summary" sheetId="1" r:id="rId1"/>
  <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
```

The `name` and `state` attributes are workbook metadata. The actual worksheet XML is in the target part resolved by `r:id`.

## Worksheet parts

A worksheet part is rooted at `<worksheet/>` and contains the required `<sheetData/>` element.

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
```

Use the workbook part to traverse worksheet parts:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

For display names or hidden state, read the workbook sheet list:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

## Cell basics

Rows are stored as `<row/>` elements and cells as `<c/>` elements. Cell references use A1 notation such as `A1` or `B2`. The `<v/>` value is either the raw value or an index into another structure, depending on the cell type.
