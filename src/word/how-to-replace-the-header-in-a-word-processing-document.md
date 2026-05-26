# Replace the header in a word processing document

Headers are stored in separate header parts and referenced from section properties. Replacing a header can mean editing the existing header part or creating a new part and updating the section reference.

Each `headerReference` points to a header part by relationship id. The relationship must be internal and use the header relationship type. The reference also has a `type` attribute that selects which header slot it fills.

## Header reference markup

```xml
<w:sectPr>
  <w:headerReference r:id="rId3" w:type="first"/>
  <w:headerReference r:id="rId5" w:type="default"/>
  <w:headerReference r:id="rId2" w:type="even"/>
</w:sectPr>
```

Sections can define first-page, even-page, and default headers. Missing references are inherited or substituted according to section settings such as `titlePg` and document settings such as even/odd headers.

## Rust workflow

Use the main document part to locate or create a header part, write the header XML, and update the default section reference:

```rust
{{#include ../../listings/word/src/lib.rs:replace_header}}
```

This example updates the default header slot. Apply the same relationship and section-property pattern for first-page or even-page headers by changing the `w:type` value and targeting the corresponding header part.

In `ooxmlsdk`, `MainDocumentPart::header_parts(&document)` traverses header parts. Generated schema types include `Header`, `HeaderReference`, and `SectionProperties`.
