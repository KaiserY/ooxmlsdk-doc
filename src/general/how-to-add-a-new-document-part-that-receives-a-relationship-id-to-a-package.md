# Add a new document part with a relationship ID

This article tracks an upstream Open XML SDK scenario where a new part is created with a caller-provided relationship ID.

`ooxmlsdk 0.6.0` supports this pattern for many typed child parts through `*_with_id` methods. For example, Custom XML parts can be added from a parent part with an explicit relationship ID by using `add_custom_xml_part_with_id`.

## Relationship IDs

Open XML relationships are identified by strings such as `rId1`. Relationship IDs are scoped to the package or part that owns the relationship file. A relationship ID that is valid under `word/document.xml` is separate from a relationship ID under another part.

When you do not care about the exact ID, prefer the auto-ID methods. They avoid collisions with existing relationships.

When interoperating with code that expects a specific relationship ID, use a `*_with_id` method and handle the error if that ID is already in use.

## Example shape

The auto-ID version is shown in [Add a new document part to a package](how-to-add-a-new-document-part-to-a-package.md). The explicit-ID form has the same shape, but passes the ID:

```rust
let custom_xml_part =
  main_part.add_custom_xml_part_with_id(&mut document, "application/xml", "rIdCustomXml")?;
```

The method creates both the part and the relationship from `main_part` to the new part.

## Current coverage

The original upstream article also creates several package-level parts from a blank document. `ooxmlsdk 0.6.0` can add many modeled parts, but this documentation currently avoids blank package creation examples until the crate exposes a high-level package creation helper. Use an existing package or template file when following the examples in this book.
