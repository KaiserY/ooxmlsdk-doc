# Create a presentation document by providing a file name

Creating a `.pptx` from scratch requires more than writing `ppt/presentation.xml`. A valid package also needs content type declarations, package relationships, presentation relationships, slide parts, and any required masters or layouts.

## Minimal package pieces

A minimal presentation package includes:

- `[Content_Types].xml`,
- `_rels/.rels` pointing to `ppt/presentation.xml`,
- `ppt/presentation.xml`,
- `ppt/_rels/presentation.xml.rels`,
- one or more `ppt/slides/slideN.xml` parts.

The listing crate builds test fixtures with that structure so every documented reader works against a real `.pptx` package.

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
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

## Creation status

`ooxmlsdk 0.6.0` can read, navigate, and save packages, and it exposes package-building primitives. This chapter does not yet publish a from-scratch presentation writer because a production example must validate:

- package and presentation relationships,
- content type overrides,
- slide id ordering,
- slide master and layout references when required,
- PowerPoint compatibility after save.

For now, use existing presentations or test fixtures as the starting point for writer examples. When this chapter gets a final writer, the code must live in `listings/presentation` and be covered by `cargo test --workspace`.
