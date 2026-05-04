# Structure of a PresentationML document

A PresentationML file is an Open Packaging Convention package. The `.pptx` file is a ZIP container whose parts are connected by relationship items. `ooxmlsdk` exposes that graph through `PresentationDocument` and generated part accessors.

## Package parts

| Package part | Root element | `ooxmlsdk` access |
|---|---|---|
| Presentation | `<p:presentation/>` | `PresentationDocument::presentation_part()` |
| Slide | `<p:sld/>` | `PresentationPart::slide_parts(&document)` |
| Slide master | `<p:sldMaster/>` | `PresentationPart::slide_master_parts(&document)` |
| Slide layout | `<p:sldLayout/>` | `SlidePart::slide_layout_part(&document)` |
| Notes slide | `<p:notes/>` | `SlidePart::notes_slide_part(&document)` |
| Notes master | `<p:notesMaster/>` | `PresentationPart::notes_master_part(&document)` |
| Handout master | `<p:handoutMaster/>` | `PresentationPart::handout_master_part(&document)` |
| Theme | `<a:theme/>` | `PresentationPart::theme_part(&document)` |
| Comments | `<p:cmLst/>` | slide comment part accessors when present |
| Comment authors | `<p:cmAuthorLst/>` | `PresentationPart::comment_authors_part(&document)` |

The exact set of parts depends on the document. A small deck can contain only the package relationship item, the presentation part, slide parts, and the required content type declarations.

## Relationship graph

The package-level relationship points to `ppt/presentation.xml`. The presentation part then owns relationships to slide parts and other top-level presentation resources.

```xml
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>
```

Inside `ppt/presentation.xml`, slide ids refer to relationship ids:

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId1"/>
  <p:sldId id="257" r:id="rId2"/>
</p:sldIdLst>
```

Those ids are resolved in `ppt/_rels/presentation.xml.rels`:

```xml
<Relationship
  Id="rId1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  Target="slides/slide1.xml"/>
```

## Reading the graph in Rust

Use the package model rather than opening ZIP paths manually:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

That pattern is the base for the other PresentationML examples: open the package, get the presentation part, then walk typed child parts or relationship iterators.

## Minimal slide part

A slide part stores only one slide. Text and drawing content are under `<p:cSld/>`.

```xml
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

For schema details, keep the XML element names separate from the Rust package API: XML stores ids and content; `ooxmlsdk` resolves ids into typed parts and relationships.
