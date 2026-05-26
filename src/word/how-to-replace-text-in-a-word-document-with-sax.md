# Replace text in a Word document using streaming XML

Streaming replacement is useful for large documents because the main document XML can be processed without loading a full model in memory.

The tradeoff is the same as other large-part workflows: a DOM-style reader is easier to query and edit, but it materializes the part; a streaming reader processes XML forward-only and can use much less memory.

## Rust workflow

Open the main document part, stream across text elements, and replace matches inside each `w:t` value:

```rust
{{#include ../../listings/word/src/lib.rs:replace_text_in_main_document}}
```

The upstream streaming pattern reads the main document part and writes updated XML to a separate in-memory stream, then replaces the original part after both reader and writer are closed. This avoids opening the same part for simultaneous read and write streams.

Text in a Word document can be split across multiple `w:t` elements. Single-word replacement is straightforward; phrase replacement needs a buffering strategy that spans runs while preserving run properties.
