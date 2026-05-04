# Open a word processing document from a stream

Some applications receive a `.docx` as bytes instead of a filesystem path. The package still has the same WordprocessingML structure: main document part, relationships, and optional supporting parts.

## Path-based baseline

The tested examples in this book open from a path:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

For stream-based workflows, use the equivalent `ooxmlsdk` package constructor for a reader that implements `Read + Seek`, such as `std::io::Cursor<Vec<u8>>`, then follow the same part traversal.

Keep stream examples in `listings/word` once added so they compile against the exact `ooxmlsdk` version used by this book.
