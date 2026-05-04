# Replace text in a Word document using streaming XML

Streaming replacement is useful for large documents because the main document XML can be processed without loading a full model in memory.

## Rust workflow

Start by locating and reading the main document part:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a streaming writer. A correct replacement must account for text split across multiple runs, preserve XML namespaces, and avoid changing unrelated field codes or hidden content.
