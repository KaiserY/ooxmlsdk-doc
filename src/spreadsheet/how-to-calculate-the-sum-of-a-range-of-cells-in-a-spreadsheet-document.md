# Calculate the sum of a range of cells in a spreadsheet

`ooxmlsdk` reads SpreadsheetML packages; it does not act as a spreadsheet calculation engine. To sum a range yourself, read the worksheet cells, select the references in the range, parse numeric values, and add them in Rust.

The upstream sample accepts a workbook path, worksheet name, first cell, last cell, and result cell. It parses row numbers from cell references, parses column names from cell references, compares columns by length and lexical order, sums cells in the rectangular range, inserts the result through the shared string table, and writes the result cell.

## Sum a range

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:sum_cell_range}}
```

The helper reads display values from the first worksheet, filters cells inside the rectangular range, parses the selected values as `f64`, and sums them in Rust. Treat values as strings at the package boundary and parse them explicitly so conversion failures are normal `Result` errors.

For a writer version, keep the read and write halves separate: first compute the numeric result, then insert or update the result cell while preserving row and cell ordering.

## Formula alternative

SpreadsheetML can store a formula and cached value:

```xml
<c r="B3">
  <f>SUM(B1:B2)</f>
  <v>42</v>
</c>
```

Editing formulas safely also involves cached values and calculation metadata. If you write a `SUM(...)` formula instead of calculating in Rust, decide whether to omit the cached `<v/>` value and let the spreadsheet application recalculate, or write a cache that matches the formula.
