# Working with notes slides

A notes slide stores speaker notes and related presentation content for a single slide. It is a separate part related from the slide part.

## Notes slide structure

The root element is `<p:notes/>`. It can contain common slide data, a color map override, and extension data. The actual note text is usually stored in text bodies under shapes.

```xml
<p:notes
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Speaker note</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>
```

Common elements:

| PresentationML element | Purpose |
|---|---|
| `<p:cSld/>` | Common notes slide data |
| `<p:clrMapOvr/>` | Color map override |
| `<p:extLst/>` | Extension data |

Important notes slide attributes include:

| Attribute | Meaning |
|---|---|
| `showMasterPhAnim` | Whether to display animations on master placeholders |
| `showMasterSp` | Whether to show shapes from the master |

`p:clrMapOvr` can use `a:masterClrMapping` to inherit the master color scheme, or `a:overrideClrMapping` to define a notes-specific mapping.

## Rust workflow

Start from the slide part and follow its notes slide relationship if present. The general traversal pattern is the same as reading slide text:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

`ooxmlsdk 0.6.0` provides the package and typed part graph. A writer for notes slides needs to update the slide relationship, notes part XML, optional notes master references, and content type declarations together; this chapter keeps that as future tested work.

When adding notes to a deck that lacks notes infrastructure, also account for the notes master part and its theme relationship. The upstream sample creates missing notes slide, notes master, and theme parts together.
