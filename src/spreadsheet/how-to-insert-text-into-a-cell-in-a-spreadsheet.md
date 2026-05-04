# Insert text into a cell in a spreadsheet

Text can be stored as an inline string or as an index into the shared string table. Excel commonly uses shared strings when saving workbooks.

## Shared string cell

```xml
<c r="A1" t="s"><v>0</v></c>
```

The `0` points into `xl/sharedStrings.xml`.

## Rust workflow

The current tested example reads shared-string-backed cells:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

This chapter does not yet publish an insertion writer. A final writer should either add an inline string or update the shared string table, then insert or update the cell in row and column order.
