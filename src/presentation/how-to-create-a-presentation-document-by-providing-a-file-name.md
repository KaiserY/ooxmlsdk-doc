# Create a presentation document by providing a file name

Creating a `.pptx` from scratch requires more than writing `ppt/presentation.xml`. A valid package also needs content type declarations, package relationships, presentation relationships, slide parts, and any required masters or layouts.

In `ooxmlsdk`, create the package with `PresentationDocument::create(PresentationDocumentType::Presentation)`, add the presentation part, add slide parts from the presentation part, write the root XML, and save the package to the file or writer your application owns.

## Minimal package pieces

A minimal presentation package includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `ppt/presentation.xml`,
- `ppt/presentation.xml`,
- `ppt/_rels/presentation.xml.rels`,
- one or more `ppt/slides/slideN.xml` parts.

This minimal writer creates a package with one slide relationship in memory:

```rust
{{#include ../../listings/presentation/src/lib.rs:create_presentation_document}}
```

## PresentationML roots

The main presentation part has a single `p:presentation` root. A usable presentation normally needs related slide, slide master, slide layout, and theme parts. The corresponding generated Rust schema types are:

| PresentationML element | Rust type |
|---|---|
| `p:presentation` | `ooxmlsdk::schemas::p::Presentation` |
| `p:sld` | `ooxmlsdk::schemas::p::Slide` |
| `p:sldMaster` | `ooxmlsdk::schemas::p::SlideMaster` |
| `p:sldLayout` | `ooxmlsdk::schemas::p::SlideLayout` |
| `a:theme` | `ooxmlsdk::schemas::a::Theme` |

The `p:presentation` root usually references slide masters, notes masters, handout masters, and slides by relationship IDs. Slide IDs are stored in `p:sldIdLst`; the relationship ID points to the slide part, while the numeric slide ID is part of the presentation markup.

## Production checks

A production presentation writer should still validate:

- package and presentation relationships,
- content type overrides,
- slide id ordering,
- slide master and layout references when required,
- PowerPoint compatibility after save.

For template workflows, `PresentationDocument::create_from_template` opens a `.potx` or `.potm` as an editable presentation package. Use `document_type()` to inspect the current type and `change_document_type(...)` when converting the package content type deliberately.
