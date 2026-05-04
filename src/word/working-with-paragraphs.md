# Working with paragraphs

Paragraphs are stored as `<w:p/>` elements in the main document body, comments, headers, footers, footnotes, and other WordprocessingML parts.

A paragraph is the basic block-level unit in WordprocessingML. It begins on a new line and can contain optional paragraph properties, inline content such as runs, and optional revision ids used by compare/merge workflows.

## Paragraph markup

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r><w:t>Heading text</w:t></w:r>
</w:p>
```

Paragraph properties are stored in `<w:pPr/>`. Text content is usually stored in runs under the paragraph.

Common paragraph properties include alignment, borders, hyphenation override, indentation, line spacing, shading, text direction, and widow/orphan control. A paragraph can exist without visible text, for example as an empty paragraph or as a cell placeholder.

## Rust workflow

Use the main document part to read paragraph text:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

If you need paragraph boundaries or properties, parse the main document XML and inspect `<w:p/>` nodes instead of flattening all text.

In ooxmlsdk 0.6.0, generated schema types include `Paragraph`, `ParagraphProperties`, `Run`, `Text`, `Justification`, `ParagraphBorders`, `Indentation`, `SpacingBetweenLines`, and `Shading`.
