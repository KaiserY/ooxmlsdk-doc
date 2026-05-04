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

Keep ownership explicit. Load a package into a document type, mutate typed parts or root elements through `&mut` bindings, then call `save` with an output writer or file path flow that your application owns. When direct XML access is unavoidable, treat it as package-level editing and revalidate the affected parts.

When you only need package read/write APIs, the default `parts` feature is enough. Enable optional features deliberately:

- Use `flat-opc` only when you need Flat OPC XML package representations.
- Use `mce` only when you want Markup Compatibility and Extensibility processing during package open and root loading.
- Use `validators` when you need validation APIs.
