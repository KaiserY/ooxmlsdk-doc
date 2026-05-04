# Get all the text in a slide in a presentation

This example reads DrawingML text runs from one slide part in a PresentationML package.

Slide text is normally stored in `a:t` elements inside shapes, placeholders, tables, SmartArt, charts, and other DrawingML containers. This simple example reads the raw slide XML and extracts `a:t` element text.

## Read slide text

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

The function:

1. Opens the presentation package lazily.
2. Gets the presentation part.
3. Selects a slide by zero-based index from `slide_parts`.
4. Reads the slide part XML.
5. Returns the text values found in `a:t` elements.

If the index is out of range, it returns an empty vector.

## Limitations

This helper is intentionally lightweight. It is suitable for simple extraction and documentation examples, but it is not a full text model for PowerPoint. Production code may need to handle tables, charts, notes, comments, ordering rules, and XML whitespace more carefully.
