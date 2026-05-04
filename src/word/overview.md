# Word processing

This section covers WordprocessingML packages (`.docx`, `.docm`, `.dotx`) with `ooxmlsdk`.

Wordprocessing packages are made of a main document part, optional styles, comments, numbering, settings, headers, footers, footnotes, endnotes, media, custom properties, and relationships between those parts. In `ooxmlsdk`, the entry point is usually `ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument`.

Use the `parts` feature, enabled by default, to open and save packages. Examples in this section are backed by tested Rust code in `listings/word`.

## In this section

- [Structure of a WordprocessingML document](structure-of-a-wordprocessingml-document.md)
- [Open a word processing document for read-only access](how-to-open-a-word-processing-document-for-read-only-access.md)
- [Retrieve comments from a word processing document](how-to-retrieve-comments-from-a-word-processing-document.md)
- [Retrieve application property values from a word processing document](how-to-retrieve-application-property-values-from-a-word-processing-document.md)
- [Extract styles from a word processing document](how-to-extract-styles-from-a-word-processing-document.md)
- [Working with paragraphs](working-with-paragraphs.md)
- [Working with runs](working-with-runs.md)
- [Working with WordprocessingML tables](working-with-wordprocessingml-tables.md)

Writer-focused chapters are being ported only when the code has a fixture in `listings/` and passes `cargo test --workspace`.

## Related sections

- [Getting started](../getting-started.md)
- [General package operations](../general/overview.md)
