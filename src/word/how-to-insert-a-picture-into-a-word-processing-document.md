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

Use the main document part as the insertion point. Add an image part, write the image bytes, and insert DrawingML that references the image relationship id.

```rust
{{#include ../../listings/word/src/lib.rs:insert_picture}}
```

When inspecting or updating existing picture relationships, preserve the relationship ids that body DrawingML references:

```rust
{{#include ../../listings/word/src/lib.rs:list_image_relationship_ids}}
```

In `ooxmlsdk`, `MainDocumentPart::image_parts(&document)` traverses existing image parts, while `related_parts_of_type::<_, ImagePart>(&document)` keeps the relationship id with each image part.
