# Open and add text to a word processing document

Adding text means editing the main document XML, usually by inserting a new paragraph or run under `<w:body/>`.

## Text markup

```xml
<w:p>
  <w:r><w:t>New text</w:t></w:r>
</w:p>
```

## Rust workflow

Read the current document text first:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a text writer. A safe implementation should parse the body XML, insert a valid paragraph or run, preserve section properties, and save the package through `ooxmlsdk`.
