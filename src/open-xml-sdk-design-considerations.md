# Design considerations

Before using `ooxmlsdk`, be clear about the level of abstraction it provides.

`ooxmlsdk` works with Office Open XML packages and generated schema types. It does not behave like Word, Excel, PowerPoint, or a full document layout engine.

## What ooxmlsdk does

- Opens and saves OOXML packages such as `.docx`, `.xlsx`, and `.pptx`.
- Exposes strongly typed package parts such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument`.
- Parses XML parts into generated Rust schema structs.
- Serializes generated schema structs back to XML.
- Preserves and round-trips package parts and relationships through the package model.

## What ooxmlsdk does not do

- It does not replace the Office application object models.
- It does not convert documents to or from formats such as HTML, PDF, XPS, or images.
- It does not calculate Word layout, paginate documents, refresh spreadsheet data, or recalculate Excel formulas.
- It does not guarantee that arbitrary generated XML is valid for every target Office version.
- It does not hide the OOXML package structure; you still need to understand parts, relationships, content types, and the relevant schema.
- It does not automatically repair files that an Office application would repair interactively.

## Rust API expectations

Use normal Rust error handling around package operations. Open, parse, and save calls can fail because input packages may be malformed, relationships may point to missing parts, XML may not match the generated schema, or the output writer may fail.

Keep ownership explicit. A document package is the sole owner of its storage and is not `Clone`. Typed Part handles are cheap to clone but remain bound to that package; pass `&package` for reads and `&mut package` for changes. Cross-package operations must copy through APIs such as `add_part_from_package` instead of resolving a source handle against a destination package.

Relationship IDs identify edges from one relationship source. They are not identities owned by target Parts, and several IDs can target the same Part. Preserve `RelatedPart` values or explicit relationship IDs whenever an edit must update both XML `r:id` references and package relationships.

Load a package into a document type, mutate raw Part data or typed root elements deliberately, then call `save` with an output writer or file path flow that your application owns. Lazy opening avoids parsing every typed root. Untouched lazy payloads can be reused, but a loaded root is serialized on save; replacing raw data unloads that root. When direct XML access is unavoidable, treat it as package-level editing and revalidate the affected parts.

Package readers and schema XML readers serve different constraints. Package constructors require `Read + Seek` for ZIP access; a generic reader is consumed during open and copied into shared in-memory archive storage. Generated root types separately offer a borrowed `from_bytes` path and a streaming `from_reader` path for `BufRead` XML sources.

Generated schema fields model OOXML values, not only primitive Rust values. Boolean-like attributes, measurements, coordinates, and percentages may use `simple_type` or `units` wrappers so unknown lexical forms, compatibility values, and unit categories can round-trip correctly.

When you only need package read/write APIs, the default `parts` feature is enough. Enable optional features deliberately:

- Use `flat-opc` only when you need Flat OPC XML package representations.
- Use `mce` only when you want Markup Compatibility and Extensibility processing during package open and root loading.
- Enable `validators` when you need structured validation diagnostics for generated schema roots or loaded package roots.
