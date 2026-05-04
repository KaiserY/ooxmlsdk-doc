# Replace the styles parts in a word processing document

Replacing styles means copying or rewriting `word/styles.xml` and keeping the main document relationship intact.

Word 2013 and later documents can contain both `styles.xml` and `stylesWithEffects.xml`. To round-trip across Word versions, replace both parts when the source and target both contain them.

## Rust workflow

Inspect the existing styles part first:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a styles replacement writer. A final implementation should handle missing styles parts, preserve relationship ids where possible, and verify that referenced style ids still exist.

The upstream workflow extracts a complete styles part from a source document, then writes that XML into the target styles part. If the requested target part does not exist, decide whether to create it or return an explicit error. Replacing styles can change document appearance immediately because paragraphs, runs, tables, and numbering reference style ids.

In ooxmlsdk 0.6.0, use `MainDocumentPart::style_definitions_part(&document)` for the normal styles part and `MainDocumentPart::styles_with_effects_part(&document)` when present.
