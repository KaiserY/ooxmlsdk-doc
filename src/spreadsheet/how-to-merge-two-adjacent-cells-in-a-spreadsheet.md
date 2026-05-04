# Merge two adjacent cells in a spreadsheet

Merged cells are stored in worksheet XML under `<mergeCells/>`. The original cell values remain in the grid; applications display the merged range as one cell.

## Merge markup

```xml
<mergeCells count="1">
  <mergeCell ref="A1:B1"/>
</mergeCells>
```

The `ref` attribute stores the merged range.

## Rust workflow

Read the worksheet XML, then insert or update `<mergeCells/>` in the correct worksheet location:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a merge writer. A safe implementation must preserve existing merged ranges and avoid overlapping merges.
