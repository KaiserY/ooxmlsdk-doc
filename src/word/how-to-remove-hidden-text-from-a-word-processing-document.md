# Remove hidden text from a word processing document

Hidden text is represented by run properties, usually `<w:vanish/>`, on runs or inherited styles.

## Hidden text markup

```xml
<w:r>
  <w:rPr><w:vanish/></w:rPr>
  <w:t>Hidden text</w:t>
</w:r>
```

## Rust workflow

Read the main document text through the package:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

This chapter does not yet publish a hidden-text remover. A correct implementation should inspect run properties and style inheritance before deleting text.
