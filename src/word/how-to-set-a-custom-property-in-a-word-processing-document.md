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

Custom properties are written through a package-level `CustomFilePropertiesPart`. This tested listing creates the part when it is absent, writes one string property, and saves the package:

```rust
{{#include ../../listings/word/src/lib.rs:set_custom_string_property}}
```

Application properties are separate from custom properties and are read through the extended properties part:

```rust
{{#include ../../listings/word/src/lib.rs:get_application_properties}}
```

The listing is deliberately narrow: it writes a single `vt:lpwstr` value. A full custom-property updater should preserve unrelated properties, allocate unique `pid` values, choose the correct value element for each type, and replace an existing property by name without duplicating it.

When updating an existing property, replacing the whole property element is often simpler than mutating the old value because the value element name encodes the property type. After insertion or replacement, keep `pid` values unique and stable for the saved part.
