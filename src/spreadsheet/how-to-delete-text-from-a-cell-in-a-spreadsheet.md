# Delete text from a cell in a spreadsheet

Deleting text from a cell is a worksheet XML edit. If the cell uses the shared string table, the cell contains a shared string index rather than the literal text.

## Cell markup

```xml
<c r="A1" t="s"><v>0</v></c>
```

Removing the text can mean removing the cell, removing its `<v/>` value, or changing the cell type depending on the desired workbook behavior.

The upstream sample also removes the shared string table entry when no other cell uses it. That is a package-wide edit: after removing one `si` entry, every later shared string index referenced by every worksheet cell must be decremented. If you do not rewrite those indexes correctly, unrelated cells can display the wrong text.

## Rust workflow

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:delete_text_from_cell}}
```

This listing clears the target cell content while leaving the worksheet row and shared string table intact. That keeps unrelated shared string indexes stable.

If your application needs to compact the shared string table, treat that as a separate workbook-wide rewrite: remove only unused `si` entries and update every affected `t="s"` cell index in every worksheet.
