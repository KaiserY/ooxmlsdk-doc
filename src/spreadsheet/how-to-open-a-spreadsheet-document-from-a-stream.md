# Open a spreadsheet document from a stream

Some applications receive an `.xlsx` as bytes instead of a filesystem path. The package still has the same SpreadsheetML structure: workbook part, worksheet parts, relationships, and optional supporting parts.

Use a stream-based open path when the caller already owns the bytes, for example from web upload handling, object storage, or another document pipeline. The callee should not assume responsibility for closing or discarding the caller's stream unless its API explicitly takes ownership.

## Open from bytes

Use any reader that implements `Read + Seek`. For in-memory bytes, wrap the `Vec<u8>` in `std::io::Cursor` and open the package from that cursor:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_from_bytes}}
```

If the stream workflow also writes to the workbook, it must follow the same invariants as path-based writing: add or update parts, relationships, and content types together, then write the package back through an explicit save path such as an owned `Cursor<Vec<u8>>` or file handle.
