# Parse and read a large spreadsheet

Large worksheets should be processed without constructing unnecessary in-memory models. SpreadsheetML stores worksheet data as XML under each worksheet part, so a scalable reader should stream or scan worksheet XML and resolve shared strings only as needed.

The key tradeoff is DOM-style parsing versus streaming parsing. A DOM-style reader is convenient because it materializes typed elements that are easy to inspect, but it loads the whole part into memory. A streaming reader reads one XML event or element at a time and is the preferred shape for worksheets that can grow to hundreds of megabytes.

## Package traversal

Start by opening the workbook, loading shared strings once, and visiting worksheet cells through a callback:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:visit_first_worksheet_cells}}
```

This listing keeps the caller's result shape outside the reader. A caller that writes each visited cell to a database or channel can avoid building a worksheet-sized `Vec` in its own code. For very large worksheets, keep this shape but replace the inner XML scan with an event parser that reads from the worksheet part stream.

## Practical notes

- Load the shared string table once if the workbook uses shared strings.
- Process rows incrementally.
- Avoid collecting all cells unless the caller needs random access.
- Treat formulas, dates, booleans, inline strings, and styled numbers as separate conversion cases.
- Keep shared string resolution separate from row iteration so callers can choose between raw cell values and formatted text.
