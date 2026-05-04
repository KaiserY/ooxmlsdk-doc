# Working with conditional formatting

Conditional formatting is stored in worksheet XML. Each `<conditionalFormatting/>` element declares the cell range it applies to and contains one or more rules.

Conditional formatting is worksheet-level markup. It can apply to ordinary cell ranges and does not require the cells to be part of a table. The `sqref` attribute stores one or more target ranges, for example `A1:A10`.

## Conditional formatting markup

```xml
<conditionalFormatting sqref="C3:C8">
  <cfRule type="top10" dxfId="1" priority="3" rank="2"/>
</conditionalFormatting>
```

Rules can express cell comparisons, top/bottom items, data bars, color scales, icon sets, duplicate values, formulas, and other conditions.

```xml
<conditionalFormatting sqref="E3:E9">
  <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
    <formula>0.5</formula>
  </cfRule>
</conditionalFormatting>
```

In ooxmlsdk 0.6.0, the generated schema types include `ConditionalFormatting`, `ConditionalFormattingRule`, `DataBar`, `ColorScale`, and `IconSet`. A `dxfId` references differential formatting in the styles part, so readers often need worksheet XML and styles XML together.

Data bars use conditional-format value objects (`cfvo`) for minimum and maximum thresholds plus a color. Color scales use two or three `cfvo` entries and matching colors. Icon sets use threshold values to decide which icon applies to each cell. Rule `priority` values are global within the worksheet, so insertion code must preserve a coherent priority order.

## Rust workflow

Read the worksheet XML through the package and inspect conditional formatting nodes:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a writer. A safe conditional formatting writer must update worksheet rules, preserve priority ordering, and ensure any referenced differential formats exist in the styles part.
