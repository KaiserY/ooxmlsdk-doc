# Working with PivotTables

PivotTables are represented by several coordinated parts. A worksheet can own one or more pivot table definition parts, while the workbook owns pivot cache definitions and those cache definitions can own cache records.

PivotTables present aggregated views of source data. The visible PivotTable cells on the worksheet are display data; the reusable source field metadata and cached records live in the pivot cache parts.

## PivotTable package model

| Part | Purpose |
|---|---|
| Pivot table definition | Layout of the PivotTable on a worksheet |
| Pivot cache definition | Fields, source data definition, shared cache items |
| Pivot cache records | Cached source records |
| Worksheet | Displayed PivotTable cells and relationship to the definition |

The pivot table definition references a pivot cache. Multiple PivotTables can use the same cache.

The pivot table definition describes which fields appear on the row axis, column axis, values area, and report filter area. The pivot cache definition describes all available source fields, including fields that a particular PivotTable is not currently using.

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

In ooxmlsdk 0.6.0, the generated part graph includes `WorkbookPart::pivot_table_cache_definition_parts`, `WorksheetPart::pivot_table_parts`, `PivotTablePart::pivot_table_cache_definition_part`, and `PivotTableCacheDefinitionPart::pivot_table_cache_records_part`. The corresponding schema types include `PivotTableDefinition`, `PivotField`, `PivotCacheDefinition`, and `PivotCacheRecords`.
