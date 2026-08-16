# Open a spreadsheet document from a stream

Some applications receive an `.xlsx` as bytes instead of a filesystem path. The package still has the same SpreadsheetML structure: workbook part, worksheet parts, relationships, and optional supporting parts.

Use a reader-based open path when the bytes come from web upload handling, object storage, or another document pipeline. `SpreadsheetDocument::new` consumes its `Read + Seek` source while opening the ZIP package and keeps shared in-memory archive data for later lazy Part reads; it does not retain the original reader.

## Open from bytes

Use any reader that implements `Read + Seek`. For in-memory bytes, wrap an owned `Vec<u8>` in `std::io::Cursor` and open the package from that cursor:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_from_bytes}}
```

The package source and output are separate. If the workflow writes back to the workbook, follow the same invariants as path-based writing: add or update parts, relationships, and content types together, then save to an explicit target such as an owned `Cursor<Vec<u8>>` or file handle.
