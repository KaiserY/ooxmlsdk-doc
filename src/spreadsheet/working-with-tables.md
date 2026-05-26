# Working with SpreadsheetML tables

A SpreadsheetML table is a logical range over worksheet cells. The worksheet stores the actual cell values; a separate table definition part stores table metadata such as name, range, columns, and filtering.

Tables organize worksheet ranges as named datasets. They can expose filter and sort controls, structured references, calculated columns, style information, and automatic expansion behavior in spreadsheet applications.

## Table markup

```xml
<table
  id="1"
  name="SalesTable"
  displayName="SalesTable"
  ref="A1:B10"
  totalsRowShown="0"
  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <autoFilter ref="A1:B10"/>
  <tableColumns count="2">
    <tableColumn id="1" name="Region"/>
    <tableColumn id="2" name="Sales"/>
  </tableColumns>
</table>
```

The worksheet points to table definition parts with `<tableParts/>` and relationships.

The table part contains metadata only; the cell data remains in the worksheet. The `ref` attribute covers the full table range, including headers. `id` and `name` must be unique across table parts, and `displayName` must also be unique across workbook defined names because formulas can reference it.

## Rust workflow

Create a table definition part, write table metadata, then add a `tableParts` reference to the worksheet:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:add_table}}
```

To keep autofilter enabled, include an `autoFilter` element, even if it has no active criteria. Table columns live under `tableColumns`, whose `count` must match the number of `tableColumn` children.

In `ooxmlsdk`, `WorksheetPart::table_definition_parts(&document)` traverses table definition parts, and generated schema types include `Table`, `TableColumn`, and `AutoFilter`.
