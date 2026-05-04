# Add a new document part to a package

This example adds a Custom XML part to the main document part of a WordprocessingML package.

Open XML packages are made of parts and relationships. A document part is not just a file in the ZIP archive; it also needs a relationship from its owning package or parent part. In `ooxmlsdk`, typed part methods create the new part and its relationship together.

## Add a Custom XML part

The example opens an existing `.docx`, gets the main document part, adds a Custom XML part, writes XML bytes into that part, and saves the package to memory.

```rust
{{#include ../../listings/getting-started/src/lib.rs:add_custom_xml_part}}
```

The key calls are:

- `WordprocessingDocument::new_from_file`, which opens the package.
- `main_document_part`, which gets the required main document part.
- `add_custom_xml_part`, which creates a child part relationship from the main document part.
- `set_data`, which replaces the part payload.
- `save`, which writes the updated package to a writer.

The function returns the updated package bytes so the caller can write them to disk, upload them, or reopen them for further processing.

## Relationship behavior

`add_custom_xml_part` chooses a new relationship ID automatically. Use the corresponding `*_with_id` methods when you need to control the relationship ID explicitly.
