# Working with the calculation chain

The calculation chain records the order in which formula cells were last calculated. It is stored in an optional workbook-level part rooted at `<calcChain/>`.

An instance of this part contains an ordered set of references to formula cells across the workbook. A package can contain no more than one calculation chain part, and its root element is `calcChain`.

## Calculation chain markup

```xml
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <c r="B2" i="1"/>
  <c r="B3" s="1"/>
  <c r="A2"/>
</calcChain>
```

Each `<c/>` entry identifies a formula cell. The `r` attribute is the cell reference, and `i` can identify the sheet index.

The generated ooxmlsdk schema types include `CalculationChain` for `<calcChain/>` and `CalculationCell` for each `<c/>` entry. The chain records the last calculation order; it does not describe the full dependency graph between formulas and input cells.

## Rust workflow

Start from the workbook part. `ooxmlsdk 0.6.0` exposes `WorkbookPart::calculation_chain_part(&document)` when the package contains one.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

When editing formulas, stale calculation chain data can be worse than no chain at all. A writer should either update the chain consistently or remove it and let the spreadsheet application rebuild calculation state.

Spreadsheet applications may rebuild calculation state in memory when opening the workbook. The order in the calculation chain is historical calculation metadata, not a command that forces every application to calculate in exactly that order.
