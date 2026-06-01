# Getting started with ooxmlsdk

`ooxmlsdk` is a Rust library for reading, writing, and round-tripping Office Open XML documents such as `.docx`, `.xlsx`, and `.pptx`. Its package API exposes generated Rust schema types, serializers, deserializers, and strongly typed package parts.

## Rust package

Add `ooxmlsdk` to your Cargo project:

```toml
[dependencies]
ooxmlsdk = "0.8.0"
```

The default feature set enables the `parts` APIs used for `.docx`, `.xlsx`, and `.pptx` packages.

The documentation examples in this book are backed by real Rust files under `listings/` and are checked with `cargo test --workspace`.

For example, this function opens a WordprocessingML package, confirms that the main document part is attached to the package, and writes the package back to memory:

```rust
{{#include ../listings/getting-started/src/lib.rs:full_example}}
```

Use `is_encrypted_office_file_path` before opening a file when your tool needs to report encrypted packages separately from malformed or unsupported packages:

```rust
{{#include ../listings/getting-started/src/lib.rs:detect_encrypted_office_file}}
```

## Crate modules

The always-available modules are:

- `common`: shared package data types and errors.
- `namespaces`: generated namespace constants and lookup helpers.
- `schemas`: generated schema structs and simple XML parsing/serialization support.
- `sdk`: package and part traits, open settings, relationship helpers, and feature-related settings.
- `simple_type`: generated simple type support.
- `units`: OOXML measure, coordinate, angle, and percentage value helpers.

Feature-gated modules are:

- `parts`: package-level APIs behind the `parts` feature.
- `validator`: optional validator APIs behind the `validators` feature.

## Feature flags

`ooxmlsdk` has a small public feature surface:

- `default`: enables `parts`; this is the recommended configuration for most users.
- `parts`: enables package-level OOXML read/write APIs such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument`.
- `flat-opc`: enables Flat OPC package read/write helpers and also enables `parts`.
- `mce`: enables Markup Compatibility and Extensibility processing and also enables `parts`.
- `validators`: exposes optional validation APIs, but this book does not publish tested validator listings for 0.8.0.

For package APIs without extra feature behavior:

```toml
[dependencies]
ooxmlsdk = { version = "0.8.0", default-features = false, features = ["parts"] }
```

For Flat OPC helpers:

```toml
[dependencies]
ooxmlsdk = { version = "0.8.0", default-features = false, features = ["flat-opc"] }
```

For MCE processing during package open and root loading:

```toml
[dependencies]
ooxmlsdk = { version = "0.8.0", default-features = false, features = ["mce"] }
```

The `validators` feature is intentionally not used by the tested listings in this book. Verify it against the exact crate version in your project before depending on it.

For optional validator APIs:

```toml
[dependencies]
ooxmlsdk = { version = "0.8.0", features = ["validators"] }
```

## Package API

With `parts` enabled, use the package type that matches the document family:

- `WordprocessingDocument` for `.docx` and related WordprocessingML packages.
- `SpreadsheetDocument` for `.xlsx` and related SpreadsheetML packages.
- `PresentationDocument` for `.pptx` and related PresentationML packages.

Common operations include creating packages with `create`, opening packages with `new`, `new_with_settings`, `new_from_file`, or `new_from_file_with_settings`; creating editable packages from templates with `create_from_template`; checking and changing the package document type with `document_type` and `change_document_type`; detecting encrypted Office files with `is_encrypted_office_file` or `is_encrypted_office_file_path`; saving with `save`; inspecting relationships and parts with methods such as `parts`, `get_all_parts`, `get_part_by_id`, and `get_parts_of_type`; and accessing well-known child parts through typed methods such as `main_document_part`, `workbook_part`, `presentation_part`, and `worksheet_parts`.

For lower-level traversal, 0.8.0 exposes related-part helpers that can preserve the relationship id alongside the typed target part. Use those when a package edit needs to update XML `r:id` references and package relationships together.

The package types also expose convenience output helpers:

- `save` writes the current package to any `Write + Seek` target.
- `copy_to` writes the package without consuming it.
- `to_package_bytes` returns an in-memory `Vec<u8>`.
- `save_as_file` writes directly to a path.

## Version coverage

`ooxmlsdk` treats Office 2007 as the compatibility baseline while generating Rust support for newer namespaces and parts present in its checked-in metadata. That includes Office 2010, 2013, 2016, 2019, 2021, Microsoft 365-era extensions, and newer upstream namespace revisions tracked by the crate.

In practice this covers later DrawingML and chart extensions, SVG and 3D-related parts, threaded comments, dynamic-array-era spreadsheet extensions, and other post-2007 additions tracked by Open XML SDK metadata.

## Schema values

The generated 0.8.0 schema API uses explicit wrappers for OOXML-specific values. Boolean-like schema attributes use types such as `BooleanValue` and `OnOffValue`, with conversion helpers such as `from_bool()` and `as_bool()`. Many lengths, coordinates, text sizes, and percentages use types from `ooxmlsdk::units`, which preserve the OOXML lexical form while still offering unit conversions.
