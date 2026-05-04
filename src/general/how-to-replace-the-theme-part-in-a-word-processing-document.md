# Replace the theme part in a word-processing document

This example replaces the WordprocessingML package's theme part with caller-provided theme XML.

A theme part contains DrawingML theme information such as color scheme, font scheme, and format scheme. In a word-processing document, the main document part can have an optional relationship to a theme part.

## Replace the theme XML

```rust
{{#include ../../listings/getting-started/src/lib.rs:replace_theme_part}}
```

The function:

1. Opens an existing `.docx`.
2. Gets the main document part.
3. Uses the existing theme part when present.
4. Adds a theme part when the main document part does not have one.
5. Writes the provided XML bytes into the theme part.
6. Saves the updated package to memory.

The `theme_xml` argument must be valid theme part XML. `set_data` writes bytes to the part; it does not prove that the XML is semantically valid for every Office version.

## Theme part relationship

The theme relationship type is:

```text
http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme
```

For WordprocessingML packages, the theme part is usually stored under `word/theme/theme1.xml`, but callers should use the part relationship rather than hard-coding the ZIP path.
