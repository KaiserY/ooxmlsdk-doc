# Insert a picture into a word processing document

Pictures require an image part, a relationship from the main document part, and DrawingML markup in the document body.

## Package model

The body markup references an image relationship id. The relationship resolves to an image part under the package.

The image bytes are stored outside `document.xml`. The body contains a run with drawing markup, and that DrawingML references the relationship id for the image part. The graphic object data element can contain application-specific graphic data under a `uri`, so the picture markup must use the expected DrawingML picture structure for Word to render it.

```xml
<w:r>
  <w:drawing>
    <!-- inline or anchored DrawingML that references r:embed="rId..." -->
  </w:drawing>
</w:r>
```

## Rust workflow

Use the main document part as the insertion point:

```rust
{{#include ../../listings/word/src/lib.rs:open_word_read_only}}
```

This chapter does not yet publish a picture writer. A final implementation should add the image part, create the relationship, insert valid drawing markup, and verify the saved document.

In ooxmlsdk 0.6.0, `MainDocumentPart::image_parts(&document)` traverses existing image parts. A writer needs package mutation support for adding a new image part and relationship, plus DrawingML generation for dimensions, non-visual properties, blip fill, and inline or anchored layout.
