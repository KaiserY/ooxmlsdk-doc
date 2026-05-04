# Insert a chart into a spreadsheet

Charts in SpreadsheetML involve drawing parts, chart parts, worksheet relationships, anchors, and often cached series data. The visible worksheet data and the chart definition are stored separately.

## Package model

A worksheet that contains a chart typically owns a drawing relationship. The drawing part owns a chart relationship. The chart part stores chart type, series, axes, and references back to worksheet ranges.

## Rust workflow

Start by locating the worksheet part:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a chart writer. A complete implementation must create or update worksheet drawing markup, drawing relationships, chart XML, chart relationships, and content type entries.
