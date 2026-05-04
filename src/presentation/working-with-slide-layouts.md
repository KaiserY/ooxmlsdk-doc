# Working with slide layouts

A slide layout is a template part that describes the placeholder and formatting structure a slide can inherit from a slide master. In PresentationML it is rooted at `<p:sldLayout/>`; in `ooxmlsdk` it is represented as a slide layout part reached through relationships.

## Slide layout structure

The layout root can contain common slide data, header/footer settings, timing, transition, color map override, and extension data.

```xml
<p:sldLayout
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  type="title">
  <p:cSld name="Title Slide">
    <p:spTree>
      <!-- placeholder and shape definitions -->
    </p:spTree>
  </p:cSld>
</p:sldLayout>
```

Important attributes include:

| Attribute | Meaning |
|---|---|
| `type` | Layout kind, such as title, blank, or title and content |
| `matchingName` | Name used when matching layouts during template changes |
| `preserve` | Whether the layout should be kept when no slide uses it |
| `showMasterSp` | Whether master slide shapes are shown |
| `showMasterPhAnim` | Whether master placeholder animations are shown |
| `userDrawn` | Whether user-drawn data should be preserved |

## Navigating layouts in Rust

A layout is normally reached from a slide part. This example opens a presentation, walks the slides, follows each slide layout relationship, and returns the layout XML.

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_layout_xml}}
```

Layout parts can also relate back to a slide master and to dependent resources such as images, charts, diagrams, embedded packages, and theme overrides. Use the generated part accessors where available, because they resolve relationship ids without hard-coding package paths.

## Editing notes

Creating a valid layout from scratch requires a coordinated slide master, layout part, relationship entries, content type overrides, and placeholder XML. Until a chapter provides a tested construction example, prefer copying an existing layout and making small schema-aware edits.
