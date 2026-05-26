# Apply a style to a paragraph in a word processing document

Paragraph styles are referenced from paragraph properties with `<w:pStyle/>`. The style definition lives in `word/styles.xml`.

Applying a style requires both a paragraph reference and a style id. The style id is the stable value used in markup; the friendly style name can be different, such as `Heading 1` for the style id `Heading1`.

## Style reference markup

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>
</w:pPr>
```

If the target paragraph has no `<w:pPr/>`, create paragraph properties before adding `<w:pStyle/>`. Paragraph properties are the place for paragraph-level formatting such as alignment, indentation, borders, line spacing, shading, text direction, and style references.

## Rust workflow

Verify the style id exists, then update the first paragraph's paragraph properties:

```rust
{{#include ../../listings/word/src/lib.rs:apply_style_to_first_paragraph}}
```

Applying a nonexistent style id does not make the document display that style. A robust writer should read the style definitions part, check by `styleId`, optionally map from style name to id, and either add the missing style or return a clear error. A styles part is optional in a minimal document, so code must handle the missing-part case explicitly.

In `ooxmlsdk`, generated schema types include `Styles`, `Style`, `ParagraphProperties`, `ParagraphStyleId`, and `StyleRunProperties`.
