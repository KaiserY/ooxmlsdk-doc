# Working with formulas

Formulas are stored in worksheet cells with `<f/>`. The cached value from the last calculation, when present, is stored in `<v/>`.

## Formula markup

```xml
<c r="A6">
  <f>SUM(A1:A5)</f>
  <v>15</v>
</c>
```

The formula text is not evaluated by `ooxmlsdk`. Spreadsheet applications may recalculate it when the workbook is opened. If you edit formula inputs or formula text, make sure cached values and calculation metadata are still appropriate for your use case.

## Rust workflow

Use the worksheet traversal pattern to find cells and inspect their XML:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_worksheet_xml}}
```

This first-pass chapter does not publish a formula writer. A final writer should parse worksheet XML, update only the intended `<f/>` and cached `<v/>` nodes, and decide whether to remove or refresh the calculation chain.
