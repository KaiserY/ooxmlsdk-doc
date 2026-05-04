# Parse and read a large spreadsheet

Large worksheets should be processed without constructing unnecessary in-memory models. SpreadsheetML stores worksheet data as XML under each worksheet part, so a scalable reader should stream or scan worksheet XML and resolve shared strings only as needed.

## Package traversal

Start by opening the workbook and locating worksheet parts:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

The listing returns XML strings because it is a compact documentation example. For very large sheets, adapt the same package traversal but process the worksheet data with a streaming XML parser.

## Practical notes

- Load the shared string table once if the workbook uses shared strings.
- Process rows incrementally.
- Avoid collecting all cells unless the caller needs random access.
- Treat formulas, dates, booleans, inline strings, and styled numbers as separate conversion cases.

A dedicated streaming example should be added under `listings/spreadsheet` before this page publishes final parser code.
