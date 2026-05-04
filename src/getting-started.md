# Getting started with ooxmlsdk

`ooxmlsdk` is a Rust library for reading, writing, and round-tripping Office Open XML documents such as `.docx`, `.xlsx`, and `.pptx`. Its package API exposes generated Rust schema types, serializers, deserializers, and strongly typed package parts.

## Rust package

Add `ooxmlsdk` to your Cargo project:

```toml
[dependencies]
ooxmlsdk = "0.6.0"
```

The default feature set enables the `parts` APIs used for `.docx`, `.xlsx`, and `.pptx` packages.

The documentation examples in this book are backed by real Rust files under `listings/` and are checked with `cargo test --workspace`.

For example, this function opens a WordprocessingML package, confirms that the main document part is attached to the package, and writes the package back to memory:

```rust
{{#include ../listings/getting-started/src/lib.rs:full_example}}
```

## Crate modules

The always-available modules are:

- `common`: shared package data types and errors.
- `schemas`: generated schema structs and simple XML parsing/serialization support.
- `sdk`: package and part traits, open settings, relationship helpers, and feature-related settings.
- `simple_type`: generated simple type support.

Feature-gated modules are:

- `parts`: package-level APIs behind the `parts` feature.
- `validator`: optional validator APIs behind the `validators` feature.

## Feature flags

`ooxmlsdk` 0.6.0 has a small public feature surface:

- `default`: enables `parts`; this is the recommended configuration for most users.
- `parts`: enables package-level OOXML read/write APIs such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument`.
- `flat-opc`: enables Flat OPC package read/write helpers and also enables `parts`.
- `mce`: enables Markup Compatibility and Extensibility processing and also enables `parts`.
- `validators`: enables optional validation APIs.

For package APIs without extra feature behavior:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["parts"] }
```

For Flat OPC helpers:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["flat-opc"] }
```

For MCE processing during package open and root loading:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["mce"] }
```

## Package API

With `parts` enabled, use the package type that matches the document family:

- `WordprocessingDocument` for `.docx` and related WordprocessingML packages.
- `SpreadsheetDocument` for `.xlsx` and related SpreadsheetML packages.
- `PresentationDocument` for `.pptx` and related PresentationML packages.

Common operations include opening packages with `new`, `new_with_settings`, `new_from_file`, or `new_from_file_with_settings`; saving with `save`; inspecting relationships and parts; and accessing well-known child parts through typed methods such as `main_document_part`, `workbook_part`, `presentation_part`, and `worksheet_parts`.

## Version coverage

`ooxmlsdk` treats Office 2007 as the compatibility baseline while generating Rust support for newer namespaces and parts present in its checked-in metadata. That includes Office 2010, 2013, 2016, 2019, 2021, and Microsoft 365-era extensions tracked by the crate.
