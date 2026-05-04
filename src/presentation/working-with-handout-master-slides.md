# Working with handout master slides

A handout master describes how printed handouts should look. It is separate from normal slide content and is related from the presentation part.

## Handout master structure

The root element is `<p:handoutMaster/>`. It can contain common slide data, a color map override, header and footer settings, extension data, and relationships to resources such as themes or images.

```xml
<p:handoutMaster
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Handout Master"/>
  <p:clrMapOvr>
    <a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
  </p:clrMapOvr>
</p:handoutMaster>
```

Common elements:

| PresentationML element | Purpose |
|---|---|
| `<p:cSld/>` | Common handout master shapes and data |
| `<p:clrMapOvr/>` | Color map override |
| `<p:hf/>` | Header and footer placeholders |
| `<p:extLst/>` | Extension data |

`p:clrMap` maps theme color names such as `bg1`, `tx1`, `accent1`, `hlink`, and `folHlink` to actual theme slots. `p:cSld` holds the common slide data for the handout master, including the shape tree and text bodies used by handout placeholders.

If a handout master is created from scratch, the package also needs a handout master part relationship and any required theme or image relationships. The upstream sample creates common slide data, a shape tree, a title placeholder shape, and a color map.

## Rust workflow

Open the package and get the presentation part. If the deck has a handout master, use `presentation_part.handout_master_part(&document)` and read the part data through the package.

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

Not every presentation contains a handout master. Treat the accessor as optional and avoid creating one from scratch until the document also has matching relationships and content type entries.
