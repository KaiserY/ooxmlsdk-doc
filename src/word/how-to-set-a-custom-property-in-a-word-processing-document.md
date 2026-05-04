# Set a custom property in a word processing document

Custom properties are stored in `docProps/custom.xml`, separate from the main document body and extended application properties.

## Custom property markup

```xml
<property name="Reviewed" pid="2">
  <vt:bool>true</vt:bool>
</property>
```

## Rust workflow

Application properties are read through a package-level part:

```rust
{{#include ../../listings/word/src/lib.rs:get_application_properties}}
```

This chapter does not yet publish a custom property writer. A final implementation should create or update the custom properties part, allocate property ids, choose the correct value type, and save package metadata.
