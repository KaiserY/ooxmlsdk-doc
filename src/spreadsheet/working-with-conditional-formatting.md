# Working with conditional formatting

Conditional formatting is stored in worksheet XML. Each `<conditionalFormatting/>` element declares the cell range it applies to and contains one or more rules.

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

## Rust workflow

Read the worksheet XML through the package and inspect conditional formatting nodes:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet publish a writer. A safe conditional formatting writer must update worksheet rules, preserve priority ordering, and ensure any referenced differential formats exist in the styles part.
