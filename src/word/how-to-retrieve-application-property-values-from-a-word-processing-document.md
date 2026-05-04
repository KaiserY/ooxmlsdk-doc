# Retrieve application property values from a word processing document

Application properties are stored in the extended file properties part, usually `docProps/app.xml`.

These are package-level extended properties, not content in the main document part. They can describe application name, page count, word count, template, company, presentation format, and other producer-maintained metadata.

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

Unlike core properties, individual extended properties can be absent when the producing application did not set them. Check the part and each requested element before reading text.

In ooxmlsdk 0.6.0, `WordprocessingDocument::extended_file_properties_part()` locates the part when present.
