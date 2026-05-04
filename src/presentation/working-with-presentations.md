# Working with presentations

The main part of a PresentationML package is `ppt/presentation.xml`. In `ooxmlsdk`, it is exposed through `PresentationDocument::presentation_part()`, and its child relationships connect the presentation to slide, slide master, notes master, handout master, theme, comment author, and other parts.

Use typed part accessors for package navigation whenever possible. They preserve the relationship model and avoid depending on ZIP paths.

## PresentationML root

The `<p:presentation/>` element stores presentation-wide properties and lists the related slides and masters. A small presentation commonly looks like this:

```xml
<p:presentation
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
    <p:sldId id="257" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
```

The `r:id` values are relationship ids. `ooxmlsdk` resolves those relationships through the package model.

Common children of `p:presentation` include:

| Element | Meaning |
|---|---|
| `p:sldMasterIdLst` | Slide masters available in the presentation |
| `p:sldIdLst` | Slides available in presentation order |
| `p:notesMasterIdLst` | Notes masters |
| `p:handoutMasterIdLst` | Handout masters |
| `p:custShowLst` | Custom slide shows with their own slide ordering |
| `p:sldSz` | Slide surface size |
| `p:notesSz` | Notes and handout surface size |
| `p:defaultTextStyle` | Default text styles |

Slide size describes the presentation slide surface; notes size describes the surface used for notes slides and handouts. Both are separate from the physical package part size.

## Common Rust workflow

Open the package, get the presentation part, then traverse child parts:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

For read-only analysis, `PackageOpenMode::Lazy` keeps startup cheap and loads part data as needed. For editing workflows, open the package, mutate supported parts or raw part data deliberately, then call `save` or `copy_to`.

## Common presentation parts

| Package concept | `ooxmlsdk` accessor |
|---|---|
| Main presentation | `presentation_part()` |
| Slides | `PresentationPart::slide_parts(&document)` |
| Slide masters | `PresentationPart::slide_master_parts(&document)` |
| Notes master | `PresentationPart::notes_master_part(&document)` |
| Handout master | `PresentationPart::handout_master_part(&document)` |
| Theme | `PresentationPart::theme_part(&document)` |
| Comment authors | `PresentationPart::comment_authors_part(&document)` |

The Rust API mirrors Open Packaging Convention relationships. Treat the XML schema element names and the package part graph as separate layers: the XML uses `r:id`, while the SDK resolves that id to a typed part or relationship.
