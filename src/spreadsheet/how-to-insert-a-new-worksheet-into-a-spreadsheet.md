# Insert a new worksheet into a spreadsheet

Inserting a worksheet requires creating a worksheet part, adding a workbook relationship, and adding a `<sheet/>` entry to the workbook's `<sheets/>` collection.

## Workbook markup

```xml
<sheets>
  <sheet name="Summary" sheetId="1" r:id="rId1"/>
  <sheet name="New Sheet" sheetId="2" r:id="rId2"/>
</sheets>
```

The new `sheetId` must be unique in the workbook, and the `r:id` must resolve to the new worksheet part.

## Rust workflow

Read the existing worksheet list before choosing a new name and id:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

This chapter does not yet publish an insertion writer. A complete example must update workbook XML, relationships, content type overrides, and save behavior together.
