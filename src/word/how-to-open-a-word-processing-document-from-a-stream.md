# Open a word processing document from a stream

Some applications receive a `.docx` as bytes instead of a filesystem path. The package still has the same WordprocessingML structure: main document part, relationships, and optional supporting parts.

Use a reader-based open path when the bytes come from an upload, object-store response, or document-processing pipeline. `WordprocessingDocument::new` consumes its `Read + Seek` source while opening the ZIP package and keeps shared in-memory archive data for later lazy Part reads; it does not retain the original reader.

## Open from bytes

Use any reader that implements `Read + Seek`. For in-memory bytes, wrap an owned `Vec<u8>` in `std::io::Cursor` and open the package from that cursor:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_from_bytes}}
```

The package source and output are separate. If the workflow writes back to the document, update parts, relationships, and content types consistently, then persist the package through an explicit output target such as an owned `Cursor<Vec<u8>>` or file handle.
