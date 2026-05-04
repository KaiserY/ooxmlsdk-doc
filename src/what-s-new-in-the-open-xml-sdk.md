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
