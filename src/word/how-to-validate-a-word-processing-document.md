# Validate a word processing document

Validation can mean different things: package relationships resolve, required parts exist, XML is well-formed, and the content follows WordprocessingML schema rules.

## Package-level checks

Start with the package graph. If `WordprocessingDocument` can open the package and retrieve the main document part, the basic OPC structure is usable.

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

Schema-level validation is broader than this first-pass chapter. A final validation example should report package errors, XML parse errors, and schema diagnostics separately.
