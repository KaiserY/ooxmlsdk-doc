# Working with presentation slides

A slide is stored as a separate PresentationML part, usually under `ppt/slides/`. In `ooxmlsdk`, slide parts are reached from the presentation part with `slide_parts(&document)`.

## Slide structure

The root element of a slide part is `<p:sld/>`. It contains common slide data, optional transition and timing data, color map overrides, and extension data.

```xml
<p:sld
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p><a:r><a:t>Slide title</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
```

Common child elements include:

| PresentationML element | Purpose |
|---|---|
| `<p:cSld/>` | Common slide data, including shapes and text |
| `<p:clrMapOvr/>` | Color map override for this slide |
| `<p:transition/>` | Transition from the previous slide |
| `<p:timing/>` | Animation and timing data |
| `<p:extLst/>` | Extension data |

## Reading slide content

For inspection tasks, read each `SlidePart` through the package and parse the text or relationships you need:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

The companion example that reads every slide uses the same traversal:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_all_slide_text}}
```

## Relationship model

Slide parts can own relationships to layouts, images, charts, embedded packages, media, comments, and hyperlinks. Prefer typed accessors such as `slide_layout_part(&document)` when the target is a known part type. For reference relationships like external hyperlinks, use relationship iterators such as `hyperlink_relationships(&document)`.

When a feature is not yet represented by a high-level helper in `ooxmlsdk`, keep the package model intact: read the part data, make a focused XML change, and save the package only after the resulting part graph is still valid.
