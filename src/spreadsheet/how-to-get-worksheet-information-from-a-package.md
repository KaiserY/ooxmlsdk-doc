# Get worksheet information from a package

Worksheet data lives in worksheet parts related from the workbook part. Use the workbook part to traverse worksheets, then read each worksheet's XML through the package.

## Read worksheet XML

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This returns one XML string per worksheet part. From there you can inspect `<sheetData/>`, dimensions, merged cells, column definitions, page setup, tables, drawings, or other worksheet-level markup.

If you need user-facing worksheet metadata, inspect the workbook `sheets` collection instead of only walking worksheet parts. Each `sheet` element carries the display `name`, workbook-local `sheetId`, relationship id, and optional visibility state. The relationship id points from the workbook part to the actual worksheet part.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

## Worksheet markup

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
```

The order returned by `worksheet_parts(&document)` follows workbook relationships. If you need sheet display names or hidden state, read the workbook `<sheets/>` list as well.
