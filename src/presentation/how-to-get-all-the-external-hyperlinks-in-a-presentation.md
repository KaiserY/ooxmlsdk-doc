# Get all the external hyperlinks in a presentation

Use `PresentationDocument` to open a `.pptx` package, iterate the slide parts, and read hyperlink relationships from each slide. In PresentationML, the visible hyperlink markup stores a relationship id such as `r:id="rLink1"`, while the actual URL is stored in the slide part relationship item.

The example opens the presentation lazily and returns only external hyperlink relationships that are referenced by a hyperlink element in slide XML.

## Read hyperlink targets

```rust
{{#include ../../listings/presentation/src/lib.rs:get_external_hyperlinks}}
```

This follows the package model instead of reading ZIP entries directly:

- `presentation_part().slide_parts(&document)` walks the slides owned by the presentation part.
- `data_as_str(&document)` reads the slide XML so the referenced relationship ids can be found.
- `hyperlink_relationships(&document)` returns the hyperlink relationships attached to that slide part.
- `relationship.target()` is the external URI target stored in the `.rels` item.

The helper returns a `Vec<String>` so callers can decide whether to print, validate, rewrite, or filter the links.

## Hyperlink relationship structure

A slide can contain hyperlink markup like this:

```xml
<a:hlinkClick r:id="rLink1"/>
```

The matching relationship item stores the URL:

```xml
<Relationship
  Id="rLink1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com/intro"
  TargetMode="External"/>
```

The relationship is scoped to the slide part that owns the markup, so the same id can appear in a different part with a different target.

When the relationship id is omitted, the hyperlink has no external relationship target. It can still point inside the same presentation through an anchor. When both an external relationship id and an anchor are present, the relationship target is the external link target for this lookup.
