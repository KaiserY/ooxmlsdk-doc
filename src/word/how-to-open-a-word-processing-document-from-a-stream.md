# Open a word processing document from a stream

Some applications receive a `.docx` as bytes instead of a filesystem path. The package still has the same WordprocessingML structure: main document part, relationships, and optional supporting parts.

Use a stream-based open path when the caller already owns an open byte source, such as an upload, object-store response, or document-processing pipeline. The callee should not close or discard a borrowed stream unless the API explicitly takes ownership.

## Path-based baseline

The tested examples in this book open from a path:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

For stream-based workflows, use the equivalent `ooxmlsdk` package constructor for a reader that implements `Read + Seek`, such as `std::io::Cursor<Vec<u8>>`, then follow the same part traversal.

Keep stream examples in `listings/word` once added so they compile against the exact `ooxmlsdk` version used by this book.

If the stream workflow writes back to the document, it must still update parts, relationships, and content types consistently, then persist the package through an explicit output path. This page remains read-oriented until a tested stream writer listing exists.
