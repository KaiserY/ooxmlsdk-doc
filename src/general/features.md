# Cargo feature flags

`ooxmlsdk 0.6.0` uses Cargo features to control optional API surface.

This page is about Rust crate features. The upstream SDK feature-collection extension model does not have a direct equivalent in `ooxmlsdk 0.6.0`.

Cargo features are compile-time switches. They choose which optional modules and dependencies are built into your binary; they are not per-package state, event hooks, or service registrations.

## Default features

Most users can use the crate with its default features:

```toml
[dependencies]
ooxmlsdk = "0.6.0"
```

The default feature set enables `parts`, which provides package-level read/write APIs such as:

- `ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument`
- `ooxmlsdk::parts::spreadsheet_document::SpreadsheetDocument`
- `ooxmlsdk::parts::presentation_document::PresentationDocument`

## `parts`

Enable `parts` when you want package APIs but want to disable other default behavior explicitly:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["parts"] }
```

This is the minimum feature for reading and writing `.docx`, `.xlsx`, and `.pptx` packages through the strongly typed package model.

## `flat-opc`

Enable `flat-opc` when you need Flat OPC XML package representations:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["flat-opc"] }
```

This feature also enables `parts`.

## `mce`

Enable `mce` when you want Markup Compatibility and Extensibility processing during package open and root loading:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["mce"] }
```

This feature also enables `parts`.

## `validators`

Enable `validators` when you need validation APIs:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", features = ["validators"] }
```

Validation support is opt-in so projects that only need package parsing, modification, and round-tripping do not need to compile validator dependencies.

## Not yet modeled

The upstream runtime feature collection includes behaviors such as package events, part events, disposal callbacks, random-number services, and automatic paragraph ID generation. Those are not exposed as `ooxmlsdk 0.6.0` APIs. When porting examples that depend on those services, write the state management explicitly in Rust or leave the page unchanged until a tested crate API exists.
