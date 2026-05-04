# Get all the text in all slides in a presentation

This example iterates over every slide part in a PresentationML package and extracts DrawingML text from each slide.

## Read all slide text

```rust
{{#include ../../listings/presentation/src/lib.rs:get_all_slide_text}}
```

The function returns one `Vec<String>` per slide. Each inner vector contains the text values found in `a:t` elements for that slide.

This keeps package traversal explicit:

- Open the package with `PresentationDocument`.
- Get the main presentation part.
- Iterate `presentation_part.slide_parts(&document)`.
- Read each slide part's XML payload.

## Ordering

This example uses the order returned by the package relationship model. For workflows that require exact presentation order, inspect the presentation root's slide ID list and resolve each relationship ID in that order.

## Limitations

The helper reads visible text in simple `a:t` elements. It does not interpret PowerPoint layout inheritance, notes pages, charts, SmartArt, comments, or accessibility metadata.
