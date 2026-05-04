// ANCHOR: full_example
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::theme_part::ThemePart;
use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode};

pub fn round_trip_word_document(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file(path)?;
  let main_part = document.main_document_part().expect("main document part");
  assert!(document.get_id_of_part(&main_part).is_some());

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: full_example

// ANCHOR: add_custom_xml_part
pub fn add_custom_xml_part(path: &Path, xml: &[u8]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file(path)?;
  let main_part = document.main_document_part()?;
  let custom_xml_part = main_part.add_custom_xml_part(&mut document, "application/xml")?;

  custom_xml_part.set_data(&mut document, xml.to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_custom_xml_part

// ANCHOR: add_custom_xml_part_with_id
pub fn add_custom_xml_part_with_id(
  path: &Path,
  relationship_id: &str,
  xml: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file(path)?;
  let main_part = document.main_document_part()?;
  let custom_xml_part =
    main_part.add_custom_xml_part_with_id(&mut document, "application/xml", relationship_id)?;

  custom_xml_part.set_data(&mut document, xml.to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_custom_xml_part_with_id

// ANCHOR: read_comments_part
pub fn read_comments_part(path: &Path) -> Result<Option<String>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file(path)?;
  let main_part = document.main_document_part()?;
  let Some(comments_part) = main_part.wordprocessing_comments_part(&document) else {
    return Ok(None);
  };

  Ok(comments_part.data_as_str(&document)?.map(str::to_owned))
}
// ANCHOR_END: read_comments_part

// ANCHOR: remove_settings_part
pub fn remove_settings_part(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file(path)?;
  let main_part = document.main_document_part()?;

  if let Some(settings_part) = main_part.document_settings_part(&document) {
    main_part.delete_part(&mut document, settings_part)?;
  }

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: remove_settings_part

// ANCHOR: copy_theme_part
pub fn copy_theme_part(
  source_path: &Path,
  target_path: &Path,
) -> Result<Option<Vec<u8>>, Box<dyn std::error::Error>> {
  let settings = OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  };
  let source = WordprocessingDocument::new_from_file_with_settings(source_path, settings)?;
  let mut target = WordprocessingDocument::new_from_file_with_settings(target_path, settings)?;

  let source_main = source.main_document_part()?;
  let target_main = target.main_document_part()?;
  let Some(source_theme) = source_main.theme_part(&source) else {
    return Ok(None);
  };
  let Some(target_theme) = target_main.theme_part(&target) else {
    return Ok(None);
  };

  let theme_data = source_theme.data_to_vec(&source).unwrap_or_default();
  target_theme.set_data(&mut target, theme_data)?;

  let mut buffer = Cursor::new(Vec::new());
  target.save(&mut buffer)?;
  Ok(Some(buffer.into_inner()))
}
// ANCHOR_END: copy_theme_part

// ANCHOR: replace_theme_part
pub fn replace_theme_part(
  path: &Path,
  theme_xml: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let settings = OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  };
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, settings)?;
  let main_part = document.main_document_part()?;

  let theme_part = match main_part.theme_part(&document) {
    Some(theme_part) => theme_part,
    None => main_part.add_new_part_auto_id::<_, ThemePart>(&mut document)?,
  };

  theme_part.set_data(&mut document, theme_xml.to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: replace_theme_part

// ANCHOR: search_and_replace_main_document
pub fn search_and_replace_main_document(
  path: &Path,
  search: &str,
  replacement: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let settings = OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  };
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, settings)?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = xml.replace(search, replacement);

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: search_and_replace_main_document

#[cfg(test)]
mod tests {
  use super::*;
  use ooxmlsdk::sdk::SdkPart;
  use std::io::Write;
  use std::sync::atomic::{AtomicUsize, Ordering};

  static FIXTURE_COUNTER: AtomicUsize = AtomicUsize::new(0);

  #[test]
  fn round_trips_wordprocessing_document() {
    let fixture = write_minimal_docx_fixture();

    let bytes = round_trip_word_document(&fixture).expect("round trip document");
    let mut reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen round-tripped docx");
    let main_part = reopened.main_document_part().expect("main document part");
    let document = main_part
      .root_element(&mut reopened)
      .expect("main document root");

    assert!(document.body.is_some());
  }

  #[test]
  fn adds_custom_xml_part() {
    let fixture = write_minimal_docx_fixture();
    let bytes =
      add_custom_xml_part(&fixture, br#"<root><value>Hello</value></root>"#).expect("add part");
    let reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen with custom xml part");
    let main_part = reopened.main_document_part().expect("main document part");

    assert_eq!(main_part.custom_xml_parts(&reopened).count(), 1);
  }

  #[test]
  fn adds_custom_xml_part_with_explicit_relationship_id() {
    let fixture = write_minimal_docx_fixture();
    let bytes = add_custom_xml_part_with_id(
      &fixture,
      "rIdCustomXml",
      br#"<root><value>Hello</value></root>"#,
    )
    .expect("add part with relationship id");
    let reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen with custom xml part");
    let main_part = reopened.main_document_part().expect("main document part");
    let custom_part = main_part
      .custom_xml_parts(&reopened)
      .next()
      .expect("custom XML part");

    assert_eq!(custom_part.relationship_id(), Some("rIdCustomXml"));
  }

  #[test]
  fn reads_comments_part() {
    let fixture = write_minimal_docx_fixture();
    let comments = read_comments_part(&fixture).expect("read comments");

    assert!(comments.expect("comments part").contains("<w:comments"));
  }

  #[test]
  fn removes_settings_part() {
    let fixture = write_minimal_docx_fixture();
    let bytes = remove_settings_part(&fixture).expect("remove settings part");
    let reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen without settings");
    let main_part = reopened.main_document_part().expect("main document part");

    assert!(main_part.document_settings_part(&reopened).is_none());
  }

  #[test]
  fn copies_theme_part() {
    let source = write_docx_fixture(true);
    let target = write_docx_fixture(true);

    let bytes = copy_theme_part(&source, &target)
      .expect("copy theme")
      .expect("theme copied");
    let reopened = WordprocessingDocument::new_with_settings(Cursor::new(bytes), lazy_settings())
      .expect("reopen copied theme");
    let main_part = reopened.main_document_part().expect("main document part");
    let theme_part = main_part.theme_part(&reopened).expect("theme part");

    assert!(
      theme_part
        .data_as_str(&reopened)
        .expect("theme xml")
        .expect("theme data")
        .contains("Office Theme")
    );
  }

  #[test]
  fn replaces_theme_part() {
    let fixture = write_minimal_docx_fixture();
    let bytes = replace_theme_part(&fixture, replacement_theme_xml()).expect("replace theme");
    let reopened = WordprocessingDocument::new_with_settings(Cursor::new(bytes), lazy_settings())
      .expect("reopen replaced theme");
    let main_part = reopened.main_document_part().expect("main document part");
    let theme_part = main_part.theme_part(&reopened).expect("theme part");

    assert!(
      theme_part
        .data_as_str(&reopened)
        .expect("theme xml")
        .expect("theme data")
        .contains("Replacement Theme")
    );
  }

  #[test]
  fn searches_and_replaces_main_document_text() {
    let fixture = write_minimal_docx_fixture();
    let bytes =
      search_and_replace_main_document(&fixture, "Hello World", "Hi Everyone").expect("replace");
    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen replaced doc");
    let main_part = reopened.main_document_part().expect("main document part");

    assert!(
      main_part
        .data_as_str(&reopened)
        .expect("document xml")
        .expect("document data")
        .contains("Hi Everyone")
    );
  }

  fn write_minimal_docx_fixture() -> std::path::PathBuf {
    write_docx_fixture(false)
  }

  fn write_docx_fixture(include_theme: bool) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-getting-started-{}-{}.docx",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    let file = std::fs::File::create(&path).expect("create fixture");
    let mut zip = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();

    zip
      .start_file("[Content_Types].xml", options)
      .expect("content types");
    let content_types = format!(
      r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  {theme_content_type}
</Types>"#,
      theme_content_type = if include_theme {
        r#"<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#
      } else {
        ""
      },
    );
    zip
      .write_all(content_types.as_bytes())
      .expect("write content types");

    zip.add_directory("_rels", options).expect("rels dir");
    zip
      .start_file("_rels/.rels", options)
      .expect("package rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .expect("write package rels");

    zip.add_directory("word", options).expect("word dir");
    zip
      .add_directory("word/_rels", options)
      .expect("word rels dir");
    zip
      .start_file("word/document.xml", options)
      .expect("document part");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>"#,
    )
    .expect("write document");

    zip
      .start_file("word/_rels/document.xml.rels", options)
      .expect("document rels");
    let document_relationships = format!(
      r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  {theme_relationship}
</Relationships>"#,
      theme_relationship = if include_theme {
        r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>"#
      } else {
        ""
      },
    );
    zip
      .write_all(document_relationships.as_bytes())
      .expect("write document rels");

    zip
      .start_file("word/comments.xml", options)
      .expect("comments part");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
      )
      .expect("write comments");

    zip
      .start_file("word/settings.xml", options)
      .expect("settings part");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
      )
      .expect("write settings");

    if include_theme {
      zip.add_directory("word/theme", options).expect("theme dir");
      zip
        .start_file("word/theme/theme1.xml", options)
        .expect("theme part");
      zip.write_all(default_theme_xml()).expect("write theme");
    }

    zip.finish().expect("finish fixture");
    path
  }

  fn default_theme_xml() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements/>
</a:theme>"#
  }

  fn replacement_theme_xml() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Replacement Theme">
  <a:themeElements/>
</a:theme>"#
  }

  fn lazy_settings() -> OpenSettings {
    OpenSettings {
      open_mode: PackageOpenMode::Lazy,
      ..Default::default()
    }
  }
}
