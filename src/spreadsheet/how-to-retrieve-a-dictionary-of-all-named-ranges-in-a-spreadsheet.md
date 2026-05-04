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

Open the workbook part and inspect `xl/workbook.xml`:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_defined_names}}
```

The function returns an ordered map from defined-name text to the workbook expression stored in the element body. If the workbook has no `<definedNames/>` collection, the map is empty.

The element body can be a plain range, a sheet-qualified range, or a formula-like expression. Do not treat every value as a rectangular cell range without validating it first.
