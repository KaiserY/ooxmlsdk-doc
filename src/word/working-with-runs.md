# Working with runs

Runs are inline text containers inside paragraphs. A run can carry formatting such as bold, italic, color, font, or field-related markup.

## Run markup

```xml
<w:r>
  <w:rPr>
    <w:b/>
  </w:rPr>
  <w:t>Bold text</w:t>
</w:r>
```

Run properties are stored in `<w:rPr/>`; visible text is stored in `<w:t/>`. Text for a sentence can be split across many runs, especially after editing in Word.

## Rust workflow

The text helper extracts `<w:t/>` values from the main document:

```rust
{{#include ../../listings/word/src/lib.rs:get_document_text}}
```

For formatting-sensitive tasks, preserve run boundaries and update only the target run or run property subtree.
