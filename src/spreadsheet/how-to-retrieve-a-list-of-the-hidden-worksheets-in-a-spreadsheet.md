# Retrieve a list of the hidden worksheets in a spreadsheet

Worksheet visibility is stored in the workbook `<sheet/>` entries. A hidden worksheet usually has `state="hidden"` or `state="veryHidden"`.

## Workbook markup

```xml
<sheets>
  <sheet name="Summary" sheetId="1" r:id="rId1"/>
  <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
```

The worksheet part itself does not carry the display name or workbook visibility state.

## Rust workflow

Read the workbook sheet list and filter entries with a hidden state. The listing below extracts worksheet names from the workbook XML; a hidden-sheet helper can extend the same traversal by reading the `state` attribute.

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

Use `worksheet_parts(&document)` only after resolving which sheet entry you want to inspect.
