# Copy part contents to a part in another package

This example copies the raw data from the theme part in one WordprocessingML package to the theme part in another package.

Theme parts are optional. The example returns `Ok(None)` when either package does not have a theme part.

## Theme parts

A Theme part stores a document theme: color scheme, font scheme, and effect scheme. In WordprocessingML, a theme affects heading colors and styles; in SpreadsheetML, it affects cell and chart formatting; in PresentationML, it affects slides, handouts, notes, and masters.

WordprocessingML and SpreadsheetML packages can have zero or one Theme part associated with the main document or workbook part. PresentationML packages can have Theme parts associated with the presentation, slide masters, notes masters, or handout masters.

## Copy the theme part

```rust
{{#include ../../listings/getting-started/src/lib.rs:copy_theme_part}}
```

The function opens the source package and target package, finds each main document part, looks up each theme part, copies the source bytes, and saves the updated target package to memory.

This is a raw part-data copy. It does not parse or validate the theme XML.

## When to copy raw part data

Raw part copying is useful when:

- You want to preserve a part exactly as stored.
- You do not need to inspect or modify the XML structure.
- The source and target part types are the same.

Use generated schema root elements instead when you need to read or modify specific XML elements or attributes.

Before running this workflow on real files, make sure the source document actually has a Theme part. If the target lacks one, create it first, as shown in [Replace the theme part in a word processing document](how-to-replace-the-theme-part-in-a-word-processing-document.md).
