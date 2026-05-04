# Retrieve application property values from a word processing document

Application properties are stored in the extended file properties part, usually `docProps/app.xml`.

## Read application properties

```rust
{{#include ../../listings/word/src/lib.rs:get_application_properties}}
```

The example reads a few common extended properties from the package-level part.

## Extended properties markup

```xml
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>ooxmlsdk-doc</Application>
  <Pages>1</Pages>
  <Words>4</Words>
</Properties>
```

Not every document contains every property. Treat this part as optional.
