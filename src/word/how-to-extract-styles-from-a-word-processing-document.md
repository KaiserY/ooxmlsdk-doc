# Extract styles from a word processing document

Styles are stored in the style definitions part, usually `word/styles.xml`. Paragraphs and runs refer to styles by id.

## Read style ids

```rust
{{#include ../../listings/word/src/lib.rs:get_style_ids}}
```

The helper follows the main document's styles relationship and extracts `w:styleId` values.

## Style markup

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/>
</w:style>
```

Styles can define paragraph, character, table, and numbering behavior. A style id is the stable value referenced from document content.
