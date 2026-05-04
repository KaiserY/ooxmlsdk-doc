# Working with the shared string table

The shared string table stores one copy of each workbook string. Worksheet cells can then store an index into that table instead of repeating the same text in every cell.

## Shared string structure

The table is rooted at `<sst/>`, and each shared string item is stored in `<si/>`.

```xml
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Region</t></si>
  <si><t>Sales</t></si>
  <si><t>North</t></si>
</sst>
```

A cell that references the first shared string stores `t="s"` and `<v>0</v>`:

```xml
<c r="A1" t="s"><v>0</v></c>
```

## Reading shared strings in Rust

The cell value example reads the workbook's shared string table, then resolves shared string indexes while reading worksheet cells:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:get_cell_values}}
```

Rich text shared strings can contain multiple runs under a single `<si/>`. A reader must concatenate or preserve those runs depending on whether it needs plain text or formatting.
