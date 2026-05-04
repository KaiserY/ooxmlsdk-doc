# Merge two adjacent cells in a spreadsheet

Merged cells are stored in worksheet XML under `<mergeCells/>`. The original cell values remain in the grid; applications display the merged range as one cell.

Before adding a merge range, make sure both endpoint cells exist. A writer normally creates any missing row or cell elements first, then records the range in `mergeCells`.

## Merge markup

```xml
<mergeCells count="1">
  <mergeCell ref="A1:B1"/>
</mergeCells>
```

The `ref` attribute stores the merged range.

The range is expressed with normal cell references such as `A1:B1`. Parse a cell reference by separating the column letters from the row digits; this is useful both for checking adjacency and for inserting any missing cells in row and column order.

## Rust workflow

Read the worksheet XML, then insert or update `<mergeCells/>` in the correct worksheet location:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a merge writer. A safe implementation must preserve existing merged ranges and avoid overlapping merges.

Only one cell's displayed content is preserved by spreadsheet applications after a merge. For left-to-right sheets, that is typically the upper-left cell in the merged range; for right-to-left sheets, applications can preserve the upper-right value. Avoid relying on values from the other cells in the range.
