# Working with PivotTables

PivotTables are represented by several coordinated parts. A worksheet can own one or more pivot table definition parts, while the workbook owns pivot cache definitions and those cache definitions can own cache records.

## PivotTable package model

| Part | Purpose |
|---|---|
| Pivot table definition | Layout of the PivotTable on a worksheet |
| Pivot cache definition | Fields, source data definition, shared cache items |
| Pivot cache records | Cached source records |
| Worksheet | Displayed PivotTable cells and relationship to the definition |

The pivot table definition references a pivot cache. Multiple PivotTables can use the same cache.

## PivotTable markup

```xml
<pivotTableDefinition name="PivotTable1" cacheId="1">
  <location ref="A3:C10" firstHeaderRow="1" firstDataRow="2"/>
  <pivotFields count="2">
    <pivotField axis="axisRow"/>
    <pivotField dataField="1"/>
  </pivotFields>
</pivotTableDefinition>
```

## Rust workflow

Use the workbook and worksheet part graph to locate pivot-related parts. `WorkbookPart` exposes pivot cache definition parts, and `WorksheetPart` exposes pivot table parts.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This first-pass page is read-oriented. Creating or editing PivotTables requires coordinated cache, worksheet, relationship, and display cell updates, so writer code should be added only with full fixture coverage.
