# Create and add a paragraph style to a word processing document

Paragraph styles are stored in the styles part with `w:type="paragraph"`.

## Style markup

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/>
  <w:pPr/>
  <w:rPr/>
</w:style>
```

## Rust workflow

Use the style extraction helper to inspect existing style ids:

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

This chapter does not yet publish a paragraph style writer. A safe implementation must preserve existing styles and update relationships and content types if the styles part is absent.
