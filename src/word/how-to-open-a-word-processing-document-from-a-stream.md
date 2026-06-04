# Open a word processing document from a stream

Some applications receive a `.docx` as bytes instead of a filesystem path. The package still has the same WordprocessingML structure: main document part, relationships, and optional supporting parts.

Use a stream-based open path when the caller already owns an open byte source, such as an upload, object-store response, or document-processing pipeline. The callee should not close or discard a borrowed stream unless the API explicitly takes ownership.

## Open from bytes

Use any reader that implements `Read + Seek`. For in-memory bytes, wrap the `Vec<u8>` in `std::io::Cursor` and open the package from that cursor:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_from_bytes}}
```

If the stream workflow writes back to the document, it must still update parts, relationships, and content types consistently, then persist the package through an explicit output path such as an owned `Cursor<Vec<u8>>` or file handle.
