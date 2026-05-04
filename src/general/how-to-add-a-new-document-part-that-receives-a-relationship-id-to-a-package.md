# Add a new document part with a relationship ID

Some package edits need a caller-provided relationship ID. `ooxmlsdk 0.6.0` supports this pattern for many typed child parts through `*_with_id` methods.

The upstream sample creates a WordprocessingML package and assigns specific relationship IDs to newly added parts. In Rust, use the same idea when the relationship ID is part of an external contract; otherwise, let the crate allocate IDs for you.

## Relationship IDs

Open XML relationships are identified by strings such as `rId1`. Relationship IDs are scoped to the package or part that owns the relationship file. A relationship ID under `word/document.xml` is separate from a relationship ID under another part.

When you do not care about the exact ID, prefer the auto-ID methods. They avoid collisions with existing relationships.

When interoperating with code that expects a specific relationship ID, use a `*_with_id` method and handle the error if that ID is already in use.

## Add a part with an explicit ID

```rust
{{#include ../../listings/getting-started/src/lib.rs:add_custom_xml_part_with_id}}
```

The method creates both the part and the relationship from `main_part` to the new part. The corresponding test reopens the package and asserts that the custom XML part uses the requested relationship ID.

## Notes

`set_data` is the Rust equivalent of filling the payload for a raw document part. The bytes you write should already be valid for the part content type; `ooxmlsdk` does not infer a schema root for arbitrary Custom XML data.

For package-level parts such as core properties, extended properties, custom properties, thumbnails, or the main document part, use the typed package methods when they exist. Use generic `add_new_part` only after checking that the generated part type has the relationship and content-type metadata you need.
