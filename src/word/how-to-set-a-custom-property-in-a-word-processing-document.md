# Set a custom property in a word processing document

Custom properties are stored in `docProps/custom.xml`, separate from the main document body and extended application properties.

Each custom property stores a name, a property id (`pid`), a fixed format id (`fmtid`), and exactly one typed value element from the document property value namespace.

## Custom property markup

```xml
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" name="Reviewed" pid="2">
  <vt:bool>true</vt:bool>
</property>
```

Common value element names include `vt:lpwstr` for strings, `vt:filetime` for timestamps, integer value elements, floating-point value elements, and `vt:bool` for booleans.

## Rust workflow

Application properties are read through a package-level part:

```rust
{{#include ../../listings/word/src/lib.rs:get_application_properties}}
```

This chapter does not yet publish a custom property writer. A final implementation should create or update the custom properties part, allocate property ids, choose the correct value type, and save package metadata.

When updating an existing property, replacing the whole property element is often simpler than mutating the old value because the value element name encodes the property type. After insertion or replacement, renumber `pid` values from 2 upward so they remain unique and stable for the saved part.

In ooxmlsdk 0.6.0, the package model includes `CustomFilePropertiesPart`; writer coverage still needs a tested listing before this page publishes mutation code.
