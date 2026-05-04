# Delete text from a cell in a spreadsheet

Deleting text from a cell is a worksheet XML edit. If the cell uses the shared string table, the cell contains a shared string index rather than the literal text.

## Cell markup

```xml
<c r="A1" t="s"><v>0</v></c>
```

Removing the text can mean removing the cell, removing its `<v/>` value, or changing the cell type depending on the desired workbook behavior.

## Rust workflow

Read cell values first so you know which cell and storage form you are editing:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

This chapter does not yet publish a deletion writer. A safe writer must update only the target cell, preserve row ordering, and decide whether unused shared string entries should remain.
