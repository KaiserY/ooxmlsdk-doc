# Add custom UI to a spreadsheet document

Custom UI is stored in package parts related from the spreadsheet document package. It is not part of worksheet cell data.

The upstream sample customizes the Excel ribbon. The custom UI XML describes a button on the Add-ins tab and points that button at a macro in the host workbook. For that scenario, the workbook is normally macro-enabled (`.xlsm`) and already contains the macro that the ribbon callback names.

## Package model

Custom UI parts commonly use Office relationship types for ribbon extensibility or user customization. A valid update needs:

- the custom UI XML part,
- the package relationship to that part,
- content type metadata,
- any images or resources referenced by the custom UI.

The ribbon extensibility part is a package-level part. If it does not exist, a writer must create it; if it already exists, a writer should update only the intended custom UI payload and preserve unrelated package relationships.

## Rust workflow

Use `SpreadsheetDocument` to open and save the package, and use package relationship APIs for custom parts when the writer is added. The normal workbook traversal remains:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

This chapter does not yet include a tested custom UI writer for `ooxmlsdk 0.6.0`.
