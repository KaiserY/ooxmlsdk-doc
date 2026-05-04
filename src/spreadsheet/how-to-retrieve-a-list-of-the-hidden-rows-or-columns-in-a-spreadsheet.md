# Retrieve a list of the hidden rows or columns in a spreadsheet

Hidden rows and columns are stored in worksheet XML. Rows use a `hidden` attribute on `<row/>`; columns use a `hidden` attribute on `<col/>`.

## Worksheet markup

```xml
<cols>
  <col min="2" max="2" hidden="1"/>
</cols>
<sheetData>
  <row r="3" hidden="1"/>
</sheetData>
```

Column definitions can cover ranges with `min` and `max`, so a single `<col/>` entry can hide more than one column.

## Rust workflow

Read the target worksheet XML and inspect `<cols/>` and `<row/>` elements:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This first-pass page does not include a dedicated hidden row/column parser yet. Add one to `listings/spreadsheet` before publishing final code.
