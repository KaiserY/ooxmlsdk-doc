# Search and replace text in a document part

This example performs a simple string replacement against the raw XML of the main document part.

This approach is intentionally limited. It can be useful for diagnostics or controlled fixtures, but it is not a robust WordprocessingML editing strategy. A search string can cross run boundaries, appear in attributes, or accidentally match XML markup.

## Simple raw XML replacement

```rust
{{#include ../../listings/getting-started/src/lib.rs:search_and_replace_main_document}}
```

The function:

1. Opens a `.docx`.
2. Reads the main document part as UTF-8 XML.
3. Applies `str::replace`.
4. Writes the updated XML bytes back to the same part.
5. Saves the package to memory.

## Safer approaches

For production document editing, prefer schema-aware traversal and updates where possible. Word text may be split across multiple runs (`w:r`) and text nodes (`w:t`), so a naive string replacement can miss visible text or corrupt the part XML.

Use this raw replacement pattern only when you control the input shape.
