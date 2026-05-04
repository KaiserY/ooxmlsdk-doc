# Remove the headers and footers from a word processing document

Headers and footers are separate parts related from the main document part and referenced from section properties.

## Section references

```xml
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId5"/>
  <w:footerReference w:type="default" r:id="rId6"/>
</w:sectPr>
```

## Rust workflow

Start from the main document part:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a remover. A complete implementation should remove section references, delete or orphan-check header/footer relationships, and preserve unrelated section properties.
