# Move a paragraph from one presentation to another

Paragraph text in slides is stored inside DrawingML text bodies, usually under `<a:p/>`. Moving a paragraph between presentations means copying XML from one slide part to another and ensuring any referenced resources still exist.

## Shape text body

A shape text body contains all visible text and visible text properties for the shape. It can contain multiple paragraphs, and each paragraph can contain multiple runs.

Its schema shape is:

| Child element | Meaning |
|---|---|
| `a:bodyPr` | Body properties |
| `a:lstStyle` | Text list styles |
| `a:p` | Text paragraphs |

## Paragraph markup

```xml
<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:r>
    <a:t>Paragraph text</a:t>
  </a:r>
</a:p>
```

If the paragraph uses only text properties, copying the `<a:p/>` subtree can be enough. If it references hyperlinks, images, embedded objects, comments, or custom extension data, the target package also needs matching relationships and parts.

The upstream sample opens both presentations, gets the first slide from each, finds the first text body in each slide, deep-clones the first source paragraph into the target text body, and replaces the source paragraph with a placeholder paragraph. The save operation must persist both packages.

## Rust workflow

Use `ooxmlsdk` to read text from the source and target slide parts:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

This chapter does not yet publish a paragraph mover because a correct writer must preserve XML namespaces, relationships, and run properties. The first tested version should start with plain text paragraphs, then expand coverage to hyperlinks and rich formatting.
