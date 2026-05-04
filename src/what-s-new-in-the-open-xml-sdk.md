# What's new in ooxmlsdk

## 0.6.0

`ooxmlsdk 0.6.0` refines the public feature model and makes the default crate configuration smaller and clearer.

### Changed

- The default feature set now enables `parts`.
- The previous `microsoft365` public feature is no longer part of the feature surface.
- Microsoft 365-era generated schema and part coverage remains included in the checked-in generated runtime.
- Optional feature behavior is now expressed through focused flags: `flat-opc`, `mce`, and `validators`.

### Added

- `flat-opc`: enables Flat OPC read/write helpers and depends on `parts`.
- `mce`: enables Markup Compatibility and Extensibility processing and depends on `parts`.

### Package and schema coverage

The upstream .NET 3.x changelog calls out package save support, part creation metadata, MCE processing, Flat OPC, validation diagnostics, and newer Office namespaces. In `ooxmlsdk 0.6.0`, those areas map to Rust APIs and Cargo features instead of a separate framework package:

- Package read/write flows use `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument` constructors such as `new`, `new_from_file`, and `new_with_settings`, then write with `save`.
- Generated part helpers carry relationship and content-type metadata through typed part APIs. Prefer methods such as typed child-part accessors, `get_part_by_id`, `get_parts_of_type`, and relationship-specific helpers over raw package editing.
- Newer generated schema and part coverage is included in the runtime. This includes post-Office 2007 namespaces and relationships for later DrawingML, spreadsheet extensions, threaded comments, SVG media, 3D model references, and Microsoft 365-era additions.
- Markup Compatibility and Extensibility behavior is enabled by `mce`; it can process known `mc:AlternateContent` and package-level compatibility flows during loading.
- Flat OPC support is enabled by `flat-opc`; it focuses on Wordprocessing Flat OPC package XML and preserves binary parts during round trips.
- Validation APIs are enabled by `validators`. The validation surface is useful for schema-oriented checks, but it is intentionally narrower than the core package read/write path.

Some upstream .NET features do not have a one-to-one Rust API. There is no runtime feature collection equivalent to `IFeatureCollection`, and unknown-element DOM editing plus markup-compatibility validator behavior are still future work in `ooxmlsdk`.

### Feature layout

Most users can depend on the crate directly:

```toml
[dependencies]
ooxmlsdk = "0.6.0"
```

For an explicit package-only build:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["parts"] }
```

For Flat OPC:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["flat-opc"] }
```

For Markup Compatibility and Extensibility processing:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["mce"] }
```

For validation APIs:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", features = ["validators"] }
```

## 0.5.x to 0.6.0 notes

If you used the default feature set, package APIs such as `WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument` remain available.

If you explicitly enabled `microsoft365`, remove that feature from your Cargo manifest. Newer generated schema coverage is part of the generated runtime rather than controlled by that public feature.

If you need Flat OPC helpers or MCE processing, enable `flat-opc` or `mce` explicitly.
