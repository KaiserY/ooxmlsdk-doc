# Open a spreadsheet document from a stream

Some applications receive an `.xlsx` as bytes instead of a filesystem path. The package still has the same SpreadsheetML structure: workbook part, worksheet parts, relationships, and optional supporting parts.

## Path-based baseline

The tested examples in this book open from a path:

```rust
{{#include ../../listings/spreadsheet/src/lib.rs:open_spreadsheet_read_only}}
```

For stream-based workflows, use the equivalent `ooxmlsdk` package constructor for a reader that implements `Read + Seek`, such as `std::io::Cursor<Vec<u8>>`, then follow the same workbook traversal.

Keep stream examples in `listings/spreadsheet` once added so they compile against the exact `ooxmlsdk` version used by this book.
