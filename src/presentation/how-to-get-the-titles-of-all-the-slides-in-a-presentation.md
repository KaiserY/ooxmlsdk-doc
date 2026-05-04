# Get the titles of all the slides in a presentation

PowerPoint slide titles are usually represented by placeholder shapes, but real-world files can vary. This simple example treats the first text value found in each slide as that slide's title.

## Read slide titles

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_titles}}
```

The function builds on the all-slide text helper and returns the first text value from each slide, or an empty string when a slide has no text.

## When to use a richer title detector

Use this helper for simple extraction and examples. For production title detection, inspect the slide's shape tree and placeholder metadata so that only title placeholders are treated as titles.
