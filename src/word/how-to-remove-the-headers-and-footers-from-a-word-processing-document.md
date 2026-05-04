# Remove the headers and footers from a word processing document

Headers and footers are separate parts related from the main document part and referenced from section properties.

Removing headers and footers requires both package and document XML changes. Deleting only the parts leaves orphaned `headerReference` or `footerReference` elements; deleting only the references can leave unused parts in the package.

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

A document can contain multiple sections, and each section can have first-page, even-page, and default header/footer references. Remove all matching `HeaderReference` and `FooterReference` elements from every `SectionProperties` node.

In ooxmlsdk 0.6.0, `MainDocumentPart::header_parts(&document)` and `MainDocumentPart::footer_parts(&document)` traverse the related parts. Generated schema types include `Header`, `Footer`, `HeaderReference`, `FooterReference`, and `SectionProperties`.
