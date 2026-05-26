# Copy a worksheet using streaming XML

Copying a worksheet with a streaming parser is useful for large sheets because the cell XML can be processed without loading a full worksheet object model.

The upstream page compares two approaches. DOM-style copying is simpler because a worksheet tree can be cloned as a whole, but it loads the full part into memory. Streaming XML copying reads and writes elements forward-only, which is better for very large worksheet parts.

## Package model

A copied worksheet needs more than copied XML. The workbook must get a new `<sheet/>` entry and relationship, and any worksheet-owned relationships may need to be copied or adjusted.

```xml
<sheet name="Copied Sheet" sheetId="3" r:id="rId3"/>
```

Even if the worksheet XML is streamed, the workbook `sheets` collection is usually small. Updating that workbook list with a structured approach is safer than trying to stream-prepend a new `sheet` entry and accidentally change sheet order.

## Rust workflow

Copy the source worksheet XML into a new worksheet part and append a workbook sheet entry:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:copy_worksheet}}
```

If the source worksheet owns drawings, tables, comments, printer settings, or other related parts, copy those relationships and rewrite references in the copied worksheet as needed. The example above covers the worksheet XML, workbook relationship, and workbook sheet list.
