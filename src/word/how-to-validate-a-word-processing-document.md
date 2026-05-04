# Validate a word processing document

Validation can mean different things: package relationships resolve, required parts exist, XML is well-formed, and the content follows WordprocessingML schema rules.

The upstream sample separates two cases: validating a normal document and validating a deliberately corrupted document that contains schema-invalid content. A useful Rust validator should make the same distinction between expected success and expected diagnostics.

## Package-level checks

Start with the package graph. If `WordprocessingDocument` can open the package and retrieve the main document part, the basic OPC structure is usable.

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

Schema-level validation is broader than this first-pass chapter. A final validation example should report package errors, XML parse errors, and schema diagnostics separately.

Opening the package and reading the main document part is only a basic package check. Schema validation should report each error with location, part, node path or element name, and a readable message. Do not mutate a document just to test validation unless the fixture is disposable; a corrupted file may fail on later opens.

ooxmlsdk 0.6.0 currently covers typed package traversal and XML parsing for these docs, but this page does not publish a full schema validator listing yet.
