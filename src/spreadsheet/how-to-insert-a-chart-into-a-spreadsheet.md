# Insert a chart into a spreadsheet

Charts in SpreadsheetML involve drawing parts, chart parts, worksheet relationships, anchors, and often cached series data. The visible worksheet data and the chart definition are stored separately.

## Package model

A worksheet that contains a chart typically owns a drawing relationship. The drawing part owns a chart relationship. The chart part stores chart type, series, axes, and references back to worksheet ranges.

The worksheet data that feeds the chart is still normal SpreadsheetML. A row (`row`) represents one worksheet row and contains zero or more cell (`c`) elements, plus optional extension data. Common row attributes include row index `r`, cell span `spans`, style index `s`, custom height, hidden state, outline level, collapsed state, and thick border flags.

A cell stores its grid address, style, type, value, metadata, and optional formula. Its children can include a formula (`f`), scalar value (`v`), inline string (`is`), and extension list. When a cell contains a shared string, `<v>` is an index into the shared string table. For numeric and most other scalar cells, `<v>` contains the value directly. Formula cells use `<f>` for the expression and `<v>` for the last calculated result. Inline strings are stored under `<is>` when a workbook does not use the shared string table for that value.

```xml
<row r="2" spans="2:12">
  <c r="C2" s="1">
    <f>PMT(B3/12,B4,-B5)</f>
    <v>672.68336574300008</v>
  </c>
  <c r="D2"><v>180</v></c>
  <c r="E2"><v>360</v></c>
</row>
```

## Rust workflow

Start by locating the worksheet part:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a chart writer. A complete implementation must create or update worksheet drawing markup, drawing relationships, chart XML, chart relationships, and content type entries.

The upstream writer sample follows this package flow: verify that the target worksheet exists, add a drawing part to the worksheet, add a chart part under that drawing, create a chart space with editing language metadata, then build a clustered column chart from keyed values. The chart definition needs category and value axes, scaling, axis positions, crossing axis references, tick label position, label alignment, label offset, and legend settings.

After the chart XML is written, the worksheet drawing positions the chart with a `TwoCellAnchor`. The anchor records the starting and ending row/column markers so Excel knows how the chart moves or resizes when worksheet rows and columns change. The graphic frame then references the chart relationship and gives the shape a name such as `Chart 1`.

With ooxmlsdk 0.6.0 this page stays structural because the documentation set does not yet have a tested chart writer listing. When adding one, keep the writer idempotent or explicitly reject existing chart anchors; the upstream sample was intended to run only once.
