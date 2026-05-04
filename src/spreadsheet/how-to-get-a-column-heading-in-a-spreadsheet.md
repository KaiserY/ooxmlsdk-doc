# Get a column heading in a spreadsheet

Column headings are ordinary cells. In many workbooks, the first row contains headings such as `Region` or `Sales`; tables can also store column names in table definition parts.

## Read heading cells

The cell value helper reads the first worksheet and resolves shared string indexes:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

For a simple worksheet where row 1 contains headings, filter the returned `(cell_reference, value)` pairs to references ending in row `1`, such as `A1` or `B1`.

## Table headings

SpreadsheetML tables store table column names in the table definition part:

```xml
<tableColumns count="2">
  <tableColumn id="1" name="Region"/>
  <tableColumn id="2" name="Sales"/>
</tableColumns>
```

Use worksheet cell values for visible grid headings; use table definition parts when you specifically need structured table column names.
