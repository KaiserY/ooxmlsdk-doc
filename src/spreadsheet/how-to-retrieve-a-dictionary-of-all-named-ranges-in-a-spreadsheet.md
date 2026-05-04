# Retrieve a dictionary of all named ranges in a spreadsheet

Named ranges are stored in the workbook part under `<definedNames/>`. Each `<definedName/>` maps a name to a reference or formula expression.

## Workbook markup

```xml
<definedNames>
  <definedName name="SalesRange">Summary!$B$2:$B$10</definedName>
</definedNames>
```

Names can be workbook-scoped or sheet-scoped. Sheet-scoped names include a `localSheetId` attribute.

## Rust workflow

Open the workbook part and inspect `xl/workbook.xml`. The worksheet list helper uses the same workbook XML read path:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

This page does not yet include a tested named-range parser. Add one to `listings/spreadsheet` before publishing final code.
