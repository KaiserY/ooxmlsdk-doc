# Open a spreadsheet document from a stream

Some applications receive an `.xlsx` as bytes instead of a filesystem path. The package still has the same SpreadsheetML structure: workbook part, worksheet parts, relationships, and optional supporting parts.

Use a stream-based open path when the caller already owns the bytes, for example from web upload handling, object storage, or another document pipeline. The callee should not assume responsibility for closing or discarding the caller's stream unless its API explicitly takes ownership.

## Path-based baseline

The tested examples in this book open from a path:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

For stream-based workflows, use the equivalent `ooxmlsdk` package constructor for a reader that implements `Read + Seek`, such as `std::io::Cursor<Vec<u8>>`, then follow the same workbook traversal.

Keep stream examples in `listings/spreadsheet` once added so they compile against the exact `ooxmlsdk` version used by this book.

If the stream workflow also writes to the workbook, it must follow the same invariants as path-based writing: add or update parts, relationships, and content types together, then write the package back through an explicit save path. This page does not publish a stream writer until that flow is covered by a listing.
