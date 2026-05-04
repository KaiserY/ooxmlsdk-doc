# Working with SpreadsheetML tables

A SpreadsheetML table is a logical range over worksheet cells. The worksheet stores the actual cell values; a separate table definition part stores table metadata such as name, range, columns, and filtering.

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

## Rust workflow

Read the worksheet XML to find table references, then use generated part accessors on `WorksheetPart` for table definition parts when present.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This chapter does not yet include a table writer. A safe implementation must update worksheet table references, table definition XML, relationships, content types, and any formulas that refer to the table display name.
