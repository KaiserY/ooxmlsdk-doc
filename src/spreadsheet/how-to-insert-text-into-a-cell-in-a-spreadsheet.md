# Insert text into a cell in a spreadsheet

Text can be stored as an inline string or as an index into the shared string table. Excel commonly uses shared strings when saving workbooks.

The upstream workflow inserts text into a new worksheet, so it combines two package operations: add a worksheet to the workbook, then insert a cell value into that worksheet. In Rust, keep those concerns separate unless the public API intentionally creates both in one call.

## Shared string cell

```xml
<c r="A1" t="s"><v>0</v></c>
```

The `0` points into `xl/sharedStrings.xml`.

For shared strings, first locate or create the shared string table. If the target text already exists, reuse its index. If it does not exist, append a new shared string item and use the new index. The cell then stores `t="s"` and the shared string index in `<v>`.

## Rust workflow

The current tested example reads shared-string-backed cells:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

This chapter does not yet publish an insertion writer. A final writer should either add an inline string or update the shared string table, then insert or update the cell in row and column order.

When inserting the cell markup, preserve worksheet ordering. Find or create the target row, then place the cell before the first existing cell whose column comes after the target column. If a cell already exists at the requested address, update that cell instead of creating a duplicate `c` element with the same reference.
