# Retrieve the values of cells in a spreadsheet

Cell values are stored in worksheet XML. Text cells often use the shared string table: the cell stores an integer index and the actual text is stored once in `xl/sharedStrings.xml`.

## Read cell values

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

The helper reads the shared string table when present, then reads the first worksheet and resolves cells with `t="s"` through that table. Numeric cells are returned from their `<v/>` value directly.

For a sheet-specific helper, first find the workbook `sheet` entry by name, use its relationship id to resolve the worksheet part, and then search that worksheet for the requested cell reference. If the sheet or cell is missing, return `None` or an explicit domain error instead of defaulting to a misleading value.

## Cell markup

```xml
<row r="1">
  <c r="A1" t="s"><v>0</v></c>
  <c r="B1"><v>42</v></c>
</row>
```

In this example, `A1` uses shared string index `0`; `B1` stores the numeric value directly. Dates, formulas, booleans, inline strings, and styled values need additional interpretation from styles and formula cache data.

Cell type drives interpretation. A missing `t` attribute usually means the value is numeric or date-like, depending on the cell style. `t="s"` means the value is a shared string index. `t="b"` stores booleans as `0` or `1`. Formula cells can contain `<f>` plus a cached `<v>` result; reading the cached value does not recalculate the formula.
