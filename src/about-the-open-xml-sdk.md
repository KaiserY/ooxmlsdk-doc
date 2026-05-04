# About ooxmlsdk

Open XML is an open standard for word-processing documents, presentations, and spreadsheets. A `.docx`, `.pptx`, or `.xlsx` file is an Open Packaging Conventions package: a ZIP archive containing XML parts, binary parts, content types, and relationship files.

`ooxmlsdk` exposes that package structure through Rust types. You can open a package, navigate to strongly typed parts such as `WordprocessingDocument`, `PresentationDocument`, and `SpreadsheetDocument`, parse root XML elements into generated schema structs, modify those structs, and save the package again.

## Package structure

An Open XML package contains:

- `[Content_Types].xml`, which maps part names and extensions to content types.
- Package-level relationships, usually in `_rels/.rels`.
- Document parts such as `word/document.xml`, `ppt/presentation.xml`, or `xl/workbook.xml`.
- Part-level relationships, such as `word/_rels/document.xml.rels`.
- Optional media, embedded objects, custom XML, comments, styles, charts, themes, and other package parts.

`ooxmlsdk` keeps these package concepts visible. The crate does not hide the file format behind a document-editor abstraction; instead, it gives you typed access to the package and schema model.

## Strongly typed Rust APIs

The runtime crate is generated from Open XML metadata. The generated surface includes:

- Package types in `ooxmlsdk::parts`, behind the `parts` feature.
- Schema structs and enums in `ooxmlsdk::schemas`.
- Shared package traits and settings in `ooxmlsdk::sdk`.
- Common package, relationship, XML, and error types in `ooxmlsdk::common`.

Most package operations return `Result<_, ooxmlsdk::common::SdkError>` or can be used with `Box<dyn std::error::Error>` in examples. Optional package relationships are represented with `Option`, and collections are exposed through Rust iterators or vectors depending on the generated schema shape.

## Feature model

In `ooxmlsdk 0.6.0`, the default feature set enables `parts`, which is what most users need for `.docx`, `.xlsx`, and `.pptx` package work.

Additional features are opt-in:

- `flat-opc`: Flat OPC package read/write helpers.
- `mce`: Markup Compatibility and Extensibility processing.
- `validators`: validation APIs.

Use `default-features = false` when you want to make the enabled surface explicit.

## Version coverage

The generated runtime uses Office 2007 as the compatibility baseline and includes newer OOXML namespaces and package relationships from later Office generations, including Office 2010, 2013, 2016, 2019, 2021, Microsoft 365-era additions, and newer upstream namespace revisions present in the checked-in metadata.

This means newer package parts and schema types are available from the Rust crate, but document validity still depends on the XML you construct and the target applications that will consume the file.
