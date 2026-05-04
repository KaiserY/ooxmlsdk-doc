# Insert a new worksheet into a spreadsheet

Inserting a worksheet requires creating a worksheet part, adding a workbook relationship, and adding a `<sheet/>` entry to the workbook's `<sheets/>` collection.

The package must already be opened for editing. In ooxmlsdk terms, the operation belongs at the workbook part boundary: create a worksheet part, initialize it with an empty worksheet and `sheetData`, add a relationship from the workbook part to that new part, then append the corresponding `sheet` entry in workbook XML.

## Workbook markup

```xml
<sheets>
  <sheet name="Summary" sheetId="1" r:id="rId1"/>
  <sheet name="New Sheet" sheetId="2" r:id="rId2"/>
</sheets>
```

The new `sheetId` must be unique in the workbook, and the `r:id` must resolve to the new worksheet part.

Choose the next id by scanning the existing workbook `sheetId` values and adding one to the maximum. The worksheet name is independent from the part path; a typical generated name is `Sheet{sheet_id}`, but production code should also avoid collisions with existing sheet names.

## Rust workflow

Read the existing worksheet list before choosing a new name and id:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:list_worksheets}}
```

This chapter does not yet publish an insertion writer. A complete example must update workbook XML, relationships, content type overrides, and save behavior together.

This page remains structural until there is a tested writer listing. The important invariant is that workbook XML, workbook relationships, and package content types are saved together; updating only `<sheets/>` leaves a workbook entry that points nowhere.
