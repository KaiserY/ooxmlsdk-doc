# Parse and read a large spreadsheet

Large worksheets should be processed without constructing unnecessary in-memory models. SpreadsheetML stores worksheet data as XML under each worksheet part, so a scalable reader should stream or scan worksheet XML and resolve shared strings only as needed.

The key tradeoff is DOM-style parsing versus streaming parsing. A DOM-style reader is convenient because it materializes typed elements that are easy to inspect, but it loads the whole part into memory. A streaming reader reads one XML event or element at a time and is the preferred shape for worksheets that can grow to hundreds of megabytes.

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

In ooxmlsdk 0.6.0, the package traversal is covered here, but the page deliberately does not present a full streaming worksheet parser until it is backed by a tested listing. When that listing is added, it should keep shared string resolution separate from row iteration so callers can choose between raw cell values and formatted text.
