# Calculate the sum of a range of cells in a spreadsheet

`ooxmlsdk` reads SpreadsheetML packages; it does not act as a spreadsheet calculation engine. To sum a range yourself, read the worksheet cells, select the references in the range, parse numeric values, and add them in Rust.

## Read values first

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

The helper returns cell references with display values from the first worksheet. For numeric calculations, treat values as strings at the package boundary and parse them explicitly.

## Formula alternative

SpreadsheetML can store a formula and cached value:

```xml
<c r="B3">
  <f>SUM(B1:B2)</f>
  <v>42</v>
</c>
```

Editing formulas safely also involves cached values and calculation metadata. A formula writer should be fixture-backed before it is documented as final code.
