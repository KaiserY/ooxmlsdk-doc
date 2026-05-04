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

Find the sheet entry by name, resolve the corresponding worksheet part, then inspect `<cols/>` and `<row/>` elements:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_hidden_rows_or_columns}}
```

Rows and columns are numbered starting at 1. Hidden rows can be collected directly from the `r` attribute of hidden `row` elements. Hidden columns need one extra step: expand every hidden `col` range from `min` through `max`, inclusive.
