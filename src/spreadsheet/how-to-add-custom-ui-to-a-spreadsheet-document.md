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

Use `SpreadsheetDocument` to open and save the package. The package-side custom UI operation is now covered by a tested listing: add a ribbon extensibility part at the package level, write the custom UI XML, and save.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:add_custom_ui_part}}
```

This covers the package relationship and part payload. A complete macro-enabled ribbon scenario still needs the workbook to contain the callback macro named by the custom UI XML, and any referenced custom UI images must be added as related parts.
