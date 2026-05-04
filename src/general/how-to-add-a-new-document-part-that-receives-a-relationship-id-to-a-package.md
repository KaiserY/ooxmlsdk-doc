# Add a new document part with a relationship ID

Some package edits need a caller-provided relationship ID. `ooxmlsdk 0.6.0` supports this pattern for many typed child parts through `*_with_id` methods.

## Relationship IDs

Open XML relationships are identified by strings such as `rId1`. Relationship IDs are scoped to the package or part that owns the relationship file. A relationship ID under `word/document.xml` is separate from a relationship ID under another part.

When you do not care about the exact ID, prefer the auto-ID methods. They avoid collisions with existing relationships.

When interoperating with code that expects a specific relationship ID, use a `*_with_id` method and handle the error if that ID is already in use.

## Add a part with an explicit ID

```rust
{{#include ../../listings/getting-started/src/lib.rs:add_custom_xml_part_with_id}}
```

The method creates both the part and the relationship from `main_part` to the new part. The corresponding test reopens the package and asserts that the custom XML part uses the requested relationship ID.

## Current coverage

This first-round page documents the explicit-ID shape for an existing package. Blank package creation examples are deferred until this guide has a tested high-level writer fixture for that workflow.
