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
