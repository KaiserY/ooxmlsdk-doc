# Introduction to markup compatibility

Markup Compatibility and Extensibility, usually shortened to MCE, is the Open XML mechanism for handling markup that may not be understood by every application or file-format version.

For example, a Word document can contain an `mc:AlternateContent` element with a newer choice and an older fallback. A consumer chooses the content it understands and ignores or removes markup that is outside its target compatibility set.

`ooxmlsdk 0.6.0` exposes MCE processing behind the `mce` Cargo feature.

## Enable MCE support

Add the feature to your manifest:

```toml
[dependencies]
ooxmlsdk = { version = "0.6.0", default-features = false, features = ["mce"] }
```

The `mce` feature also enables `parts`.

## Open settings

MCE behavior is configured through `OpenSettings`:

```rust
use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::sdk::{
  FileFormatVersion, MarkupCompatibilityProcessMode, MarkupCompatibilityProcessSettings,
  OpenSettings,
};

let settings = OpenSettings {
  markup_compatibility_process_settings: MarkupCompatibilityProcessSettings {
    process_mode: MarkupCompatibilityProcessMode::ProcessAllParts,
    target_file_format_version: FileFormatVersion::Office2016,
  },
  ..Default::default()
};

let document = WordprocessingDocument::new_from_file_with_settings("input.docx", settings)?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

The available process modes are:

- `NoProcess`: do not process MCE markup. This is the default.
- `ProcessLoadedPartsOnly`: process MCE markup for loaded parts.
- `ProcessAllParts`: process all parts. This forces eager root loading because every part must be considered.

`ProcessLoadedPartsOnly` and `ProcessAllParts` are only available when the `mce` feature is enabled.

## Target file format version

`MarkupCompatibilityProcessSettings` also includes `target_file_format_version`. This value tells the processor which Office-era namespaces should be treated as understood.

Available values include:

- `Office2007`
- `Office2010`
- `Office2013`
- `Office2016`
- `Office2019`
- `Office2021`
- `Microsoft365`

The default target is `Office2007`.

## Saving after processing

MCE processing changes the loaded root elements. If you save a package after processing, the saved package reflects the processed content. Use `NoProcess` when you want to inspect or round-trip MCE markup without filtering it.
