# Calculate the sum of a range of cells in a spreadsheet

`ooxmlsdk` reads SpreadsheetML packages; it does not act as a spreadsheet calculation engine. To sum a range yourself, read the worksheet cells, select the references in the range, parse numeric values, and add them in Rust.

The upstream sample accepts a workbook path, worksheet name, first cell, last cell, and result cell. It parses row numbers from cell references, parses column names from cell references, compares columns by length and lexical order, sums cells in the rectangular range, inserts the result through the shared string table, and writes the result cell.

## Read values first

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

The helper returns cell references with display values from the first worksheet. For numeric calculations, treat values as strings at the package boundary and parse them explicitly.

For a production Rust version, split the workflow into two parts: a read-only range scanner that returns typed numeric values, and a writer that inserts or updates the result cell while preserving row and cell ordering.

## Formula alternative

SpreadsheetML can store a formula and cached value:

```xml
<c r="B3">
  <f>SUM(B1:B2)</f>
  <v>42</v>
</c>
```

Editing formulas safely also involves cached values and calculation metadata. A formula writer should be fixture-backed before it is documented as final code.
