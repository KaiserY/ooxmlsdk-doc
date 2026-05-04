// ANCHOR: mce_open_settings
use ooxmlsdk::sdk::{
  FileFormatVersion, MarkupCompatibilityProcessMode, MarkupCompatibilityProcessSettings,
  OpenSettings,
};

pub fn process_all_parts_for_office_2007() -> OpenSettings {
  OpenSettings {
    markup_compatibility_process_settings: MarkupCompatibilityProcessSettings {
      process_mode: MarkupCompatibilityProcessMode::ProcessAllParts,
      target_file_format_version: FileFormatVersion::Office2007,
    },
    ..Default::default()
  }
}
// ANCHOR_END: mce_open_settings

#[cfg(test)]
mod tests {
  use super::*;

  #[test]
  fn builds_mce_open_settings() {
    let settings = process_all_parts_for_office_2007();

    assert_eq!(
      settings.markup_compatibility_process_settings.process_mode,
      MarkupCompatibilityProcessMode::ProcessAllParts
    );
    assert_eq!(
      settings
        .markup_compatibility_process_settings
        .target_file_format_version,
      FileFormatVersion::Office2007
    );
  }
}
