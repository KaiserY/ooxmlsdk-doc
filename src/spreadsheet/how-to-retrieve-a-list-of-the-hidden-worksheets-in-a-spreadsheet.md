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

Read the workbook sheet list and filter entries with a hidden state:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_hidden_worksheets}}
```

The workbook XML is the source of truth for this query. `worksheet_parts(&document)` gives access to worksheet content, but hidden state is part of the workbook `sheet` entry. Treat both `hidden` and `veryHidden` as hidden worksheets; the latter is not normally exposed through the worksheet tab UI.
