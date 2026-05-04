# Apply a style to a paragraph in a word processing document

Paragraph styles are referenced from paragraph properties with `<w:pStyle/>`. The style definition lives in `word/styles.xml`.

## Style reference markup

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>
</w:pPr>
```

## Rust workflow

Read available style ids before applying one:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a style application writer. A complete implementation should verify the style exists, locate the target paragraph, and update only its `<w:pPr/>`.
