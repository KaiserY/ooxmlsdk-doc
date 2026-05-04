# Working with slide masters

A slide master stores shared formatting and placeholder structure for a group of slide layouts. In PresentationML, a slide master is rooted at `<p:sldMaster/>` and is related from the main presentation part.

## Slide master structure

Common child elements include:

| PresentationML element | Purpose |
|---|---|
| `<p:cSld/>` | Common slide data, including master shapes |
| `<p:clrMap/>` | Theme color mapping |
| `<p:sldLayoutIdLst/>` | Relationships to slide layouts |
| `<p:txStyles/>` | Default title, body, and other text styles |
| `<p:hf/>` | Header and footer defaults |
| `<p:timing/>` | Master timing data |
| `<p:transition/>` | Master transition data |
| `<p:extLst/>` | Extension data |

```xml
<p:sldMaster
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld name="Office Theme"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles/>
</p:sldMaster>
```

## Navigating masters in Rust

Start at the presentation part and use the generated accessor:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

For master-specific traversal, use `presentation_part.slide_master_parts(&document)`. From a slide master, follow layout relationships with the generated slide layout part accessors where available.

## Editing notes

Slide masters, slide layouts, and slides form a coordinated graph. A valid edit may require updating:

- the presentation part's master list,
- the slide master's layout id list,
- relationship items,
- content type overrides,
- the affected slide layout and slide XML.

Until a tested construction example exists in this book, prefer inspecting existing master/layout parts or copying a known-good package structure before making focused XML changes.
