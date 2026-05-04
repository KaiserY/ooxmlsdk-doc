# Move a paragraph from one presentation to another

Paragraph text in slides is stored inside DrawingML text bodies, usually under `<a:p/>`. Moving a paragraph between presentations means copying XML from one slide part to another and ensuring any referenced resources still exist.

## Paragraph markup

```xml
<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:r>
    <a:t>Paragraph text</a:t>
  </a:r>
</a:p>
```

If the paragraph uses only text properties, copying the `<a:p/>` subtree can be enough. If it references hyperlinks, images, embedded objects, comments, or custom extension data, the target package also needs matching relationships and parts.

## Rust workflow

Use `ooxmlsdk` to read text from the source and target slide parts:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

This chapter does not yet publish a paragraph mover because a correct writer must preserve XML namespaces, relationships, and run properties. The first tested version should start with plain text paragraphs, then expand coverage to hyperlinks and rich formatting.
