// ANCHOR: open_word_read_only
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::image_part::ImagePart;
use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode, SdkPart, WordprocessingDocumentType};

pub fn open_word_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();

  Ok(xml.matches("<w:p").count())
}
// ANCHOR_END: open_word_read_only

// ANCHOR: create_word_document
pub fn create_word_document(text: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::create(WordprocessingDocumentType::Document);
  let main_part = document.add_main_document_part()?;
  let escaped_text = escape_xml_text(text);
  let xml = format!(
    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{escaped_text}</w:t></w:r></w:p></w:body></w:document>"#
  );

  main_part.set_data(&mut document, xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: create_word_document

// ANCHOR: get_document_text
pub fn get_document_text(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_text_values(xml))
}
// ANCHOR_END: get_document_text

// ANCHOR: get_comments
pub fn get_comments(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let Some(comments_part) = main_part.wordprocessing_comments_part(&document) else {
    return Ok(Vec::new());
  };
  let xml = comments_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_text_values(xml))
}
// ANCHOR_END: get_comments

// ANCHOR: get_style_ids
pub fn get_style_ids(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let Some(styles_part) = main_part.style_definitions_part(&document) else {
    return Ok(Vec::new());
  };
  let xml = styles_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_style_ids(xml))
}
// ANCHOR_END: get_style_ids

// ANCHOR: get_application_properties
pub fn get_application_properties(
  path: &Path,
) -> Result<Vec<(String, String)>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let Some(app_part) = document.extended_file_properties_part() else {
    return Ok(Vec::new());
  };
  let xml = app_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_known_app_properties(xml))
}
// ANCHOR_END: get_application_properties

// ANCHOR: add_image_part
pub fn add_image_part(
  path: &Path,
  relationship_id: &str,
  content_type: &str,
  image_bytes: &[u8],
) -> Result<(Vec<u8>, String), Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let image_part =
    main_part.add_image_part_with_id(&mut document, content_type.to_string(), relationship_id)?;

  image_part.set_data(&mut document, image_bytes.to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok((buffer.into_inner(), relationship_id.to_string()))
}
// ANCHOR_END: add_image_part

// ANCHOR: list_image_relationship_ids
pub fn list_image_relationship_ids(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;

  Ok(
    main_part
      .related_parts_of_type::<_, ImagePart>(&document)
      .map(|related| related.relationship_id().to_string())
      .collect(),
  )
}
// ANCHOR_END: list_image_relationship_ids

// ANCHOR: convert_docm_to_docx
pub fn convert_docm_to_docx(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;

  if let Some(vba_part) = main_part.vba_project_part(&document) {
    main_part.delete_part(&mut document, vba_part)?;
  }
  document.change_document_type(WordprocessingDocumentType::Document)?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: convert_docm_to_docx

// ANCHOR: set_custom_string_property
pub fn set_custom_string_property(
  path: &Path,
  name: &str,
  value: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let custom_properties_part = match document.custom_file_properties_part() {
    Some(part) => part,
    None => document.add_custom_file_properties_part()?,
  };
  let name = escape_xml_text(name);
  let value = escape_xml_text(value);
  let xml = format!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="{name}">
    <vt:lpwstr>{value}</vt:lpwstr>
  </property>
</Properties>"#
  );

  custom_properties_part.set_data(&mut document, xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: set_custom_string_property

// ANCHOR: add_paragraph_text
pub fn add_paragraph_text(path: &Path, text: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let paragraph = paragraph_xml(text);
  let updated_xml = insert_before_section_properties(xml, &paragraph)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_paragraph_text

// ANCHOR: insert_table
pub fn insert_table(path: &Path, rows: &[&[&str]]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let table = table_xml(rows);
  let updated_xml = insert_before_section_properties(xml, &table)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_table

// ANCHOR: change_text_in_first_table_cell
pub fn change_text_in_first_table_cell(
  path: &Path,
  new_text: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = replace_first_table_cell_text(xml, new_text)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: change_text_in_first_table_cell

// ANCHOR: set_first_run_font
pub fn set_first_run_font(
  path: &Path,
  font_name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = set_first_run_fonts(xml, font_name)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: set_first_run_font

fn lazy_settings() -> OpenSettings {
  OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  }
}

fn escape_xml_text(value: &str) -> String {
  value
    .replace('&', "&amp;")
    .replace('<', "&lt;")
    .replace('>', "&gt;")
}

fn escape_xml_attr(value: &str) -> String {
  escape_xml_text(value).replace('"', "&quot;")
}

fn paragraph_xml(text: &str) -> String {
  let text = escape_xml_text(text);
  format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn table_xml(rows: &[&[&str]]) -> String {
  let mut table = String::from("<w:tbl>");
  for row in rows {
    table.push_str("<w:tr>");
    for cell in *row {
      table.push_str("<w:tc>");
      table.push_str(&paragraph_xml(cell));
      table.push_str("</w:tc>");
    }
    table.push_str("</w:tr>");
  }
  table.push_str("</w:tbl>");
  table
}

fn insert_before_section_properties(
  document_xml: &str,
  element_xml: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let insert_at = document_xml
    .find("<w:sectPr")
    .or_else(|| document_xml.find("</w:body>"))
    .ok_or("document body not found")?;
  let mut updated = String::with_capacity(document_xml.len() + element_xml.len());
  updated.push_str(&document_xml[..insert_at]);
  updated.push_str(element_xml);
  updated.push_str(&document_xml[insert_at..]);
  Ok(updated)
}

fn replace_first_table_cell_text(
  document_xml: &str,
  new_text: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let table_start = document_xml.find("<w:tbl").ok_or("table not found")?;
  let table_end = document_xml[table_start..]
    .find("</w:tbl>")
    .map(|index| table_start + index + "</w:tbl>".len())
    .ok_or("table end not found")?;
  let table_xml = &document_xml[table_start..table_end];
  let text_start = find_text_start(table_xml).ok_or("table text not found")?;
  let text_open_end = table_xml[text_start..]
    .find('>')
    .map(|index| text_start + index + 1)
    .ok_or("table text start not found")?;
  let text_end = table_xml[text_open_end..]
    .find("</w:t>")
    .map(|index| text_open_end + index)
    .ok_or("table text end not found")?;
  let new_text = escape_xml_text(new_text);

  let absolute_text_open_end = table_start + text_open_end;
  let absolute_text_end = table_start + text_end;
  let mut updated = String::with_capacity(document_xml.len() + new_text.len());
  updated.push_str(&document_xml[..absolute_text_open_end]);
  updated.push_str(&new_text);
  updated.push_str(&document_xml[absolute_text_end..]);
  Ok(updated)
}

fn set_first_run_fonts(
  document_xml: &str,
  font_name: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let run_start = document_xml.find("<w:r>").ok_or("run not found")?;
  let run_open_end = run_start + "<w:r>".len();
  let run_end = document_xml[run_open_end..]
    .find("</w:r>")
    .map(|index| run_open_end + index)
    .ok_or("run end not found")?;
  let run_inner = &document_xml[run_open_end..run_end];
  let font_name = escape_xml_attr(font_name);
  let run_fonts = format!(r#"<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>"#);

  let updated_run_inner = if let Some(properties_start) = run_inner.find("<w:rPr>") {
    let insert_at = properties_start + "<w:rPr>".len();
    let mut value = String::with_capacity(run_inner.len() + run_fonts.len());
    value.push_str(&run_inner[..insert_at]);
    value.push_str(&run_fonts);
    value.push_str(&run_inner[insert_at..]);
    value
  } else {
    format!("<w:rPr>{run_fonts}</w:rPr>{run_inner}")
  };

  let mut updated = String::with_capacity(document_xml.len() + run_fonts.len());
  updated.push_str(&document_xml[..run_open_end]);
  updated.push_str(&updated_run_inner);
  updated.push_str(&document_xml[run_end..]);
  Ok(updated)
}

fn extract_text_values(xml: &str) -> Vec<String> {
  let mut values = Vec::new();
  let mut rest = xml;

  while let Some(start) = find_text_start(rest) {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    rest = &rest[tag_end + 1..];
    let Some(end) = rest.find("</w:t>") else {
      break;
    };
    values.push(decode_minimal_xml_text(&rest[..end]));
    rest = &rest[end + "</w:t>".len()..];
  }

  values
}

fn find_text_start(xml: &str) -> Option<usize> {
  match (xml.find("<w:t>"), xml.find("<w:t ")) {
    (Some(left), Some(right)) => Some(left.min(right)),
    (Some(left), None) => Some(left),
    (None, Some(right)) => Some(right),
    (None, None) => None,
  }
}

fn extract_style_ids(xml: &str) -> Vec<String> {
  let mut ids = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<w:style ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if let Some(id) = extract_attr(tag, "w:styleId") {
      ids.push(id.to_string());
    }
    rest = &rest[tag_end + 1..];
  }

  ids
}

fn extract_known_app_properties(xml: &str) -> Vec<(String, String)> {
  ["Application", "Pages", "Words"]
    .into_iter()
    .filter_map(|name| extract_element_text(xml, name).map(|value| (name.to_string(), value)))
    .collect()
}

fn extract_element_text(xml: &str, name: &str) -> Option<String> {
  let open = format!("<{name}>");
  let close = format!("</{name}>");
  let start = xml.find(&open)? + open.len();
  let end = xml[start..].find(&close)?;
  Some(decode_minimal_xml_text(&xml[start..start + end]))
}

fn extract_attr<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
  let pattern = format!(r#"{name}=""#);
  let start = tag.find(&pattern)? + pattern.len();
  let end = tag[start..].find('"')?;
  Some(&tag[start..start + end])
}

fn decode_minimal_xml_text(text: &str) -> String {
  text
    .replace("&lt;", "<")
    .replace("&gt;", ">")
    .replace("&quot;", "\"")
    .replace("&apos;", "'")
    .replace("&amp;", "&")
}

#[cfg(test)]
mod tests {
  use super::*;
  use std::io::Write;
  use std::sync::atomic::{AtomicUsize, Ordering};

  static FIXTURE_COUNTER: AtomicUsize = AtomicUsize::new(0);

  #[test]
  fn opens_word_read_only_and_counts_paragraphs() {
    let fixture = write_word_fixture();

    let count = open_word_read_only(&fixture).expect("open document");

    assert_eq!(count, 3);
  }

  #[test]
  fn creates_word_document() {
    let bytes = create_word_document("A&B").expect("create document");
    let document =
      WordprocessingDocument::new(std::io::Cursor::new(bytes)).expect("reopen document");
    let main_part = document.main_document_part().expect("main document part");
    let xml = main_part
      .data_as_str(&document)
      .expect("main document XML")
      .expect("main document bytes");

    assert!(xml.contains("<w:t>A&amp;B</w:t>"));
  }

  #[test]
  fn gets_document_text() {
    let fixture = write_word_fixture();

    let text = get_document_text(&fixture).expect("document text");

    assert_eq!(text, vec!["Hello", "from WordprocessingML", "Cell text"]);
  }

  #[test]
  fn gets_comments() {
    let fixture = write_word_fixture();

    let comments = get_comments(&fixture).expect("comments");

    assert_eq!(comments, vec!["Review this paragraph"]);
  }

  #[test]
  fn gets_style_ids() {
    let fixture = write_word_fixture();

    let styles = get_style_ids(&fixture).expect("styles");

    assert_eq!(styles, vec!["Normal", "Heading1"]);
  }

  #[test]
  fn gets_application_properties() {
    let fixture = write_word_fixture();

    let props = get_application_properties(&fixture).expect("app properties");

    assert_eq!(
      props,
      vec![
        ("Application".to_string(), "ooxmlsdk-doc".to_string()),
        ("Pages".to_string(), "1".to_string()),
        ("Words".to_string(), "4".to_string())
      ]
    );
  }

  #[test]
  fn adds_image_part_with_relationship_id() {
    let fixture = write_word_fixture();
    let image_bytes = b"\x89PNG\r\n\x1a\nimage bytes";

    let (bytes, relationship_id) =
      add_image_part(&fixture, "rIdImage1", "image/png", image_bytes).expect("add image part");

    assert_eq!(relationship_id, "rIdImage1");
    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen with image");
    let main_part = reopened.main_document_part().expect("main document part");
    let image_part = main_part
      .related_parts_of_type::<_, ImagePart>(&reopened)
      .find(|related| related.relationship_id() == "rIdImage1")
      .map(|related| related.into_part())
      .expect("image part by relationship id");

    assert_eq!(image_part.content_type(&reopened), Some("image/png"));
    assert_eq!(image_part.data(&reopened), Some(image_bytes.as_slice()));
  }

  #[test]
  fn lists_image_relationship_ids() {
    let fixture = write_word_fixture();
    let image_bytes = b"\x89PNG\r\n\x1a\nimage bytes";
    let (bytes, _) =
      add_image_part(&fixture, "rIdImage2", "image/png", image_bytes).expect("add image part");
    let path = write_bytes_fixture("docx", bytes);

    let ids = list_image_relationship_ids(&path).expect("image relationship ids");

    assert_eq!(ids, vec!["rIdImage2"]);
  }

  #[test]
  fn converts_docm_to_docx_package() {
    let fixture = write_macro_enabled_word_fixture();

    let bytes = convert_docm_to_docx(&fixture).expect("convert docm");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen converted docx");
    let main_part = reopened.main_document_part().expect("main document part");
    assert_eq!(
      reopened.document_type(),
      WordprocessingDocumentType::Document
    );
    assert!(main_part.vba_project_part(&reopened).is_none());
  }

  #[test]
  fn sets_custom_string_property() {
    let fixture = write_word_fixture();

    let bytes =
      set_custom_string_property(&fixture, "Reviewed", "yes & checked").expect("set property");

    let reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen with custom property");
    let custom_properties_part = reopened
      .custom_file_properties_part()
      .expect("custom properties part");
    let xml = custom_properties_part
      .data_as_str(&reopened)
      .expect("custom properties XML")
      .expect("custom properties data");

    assert!(xml.contains(r#"name="Reviewed""#));
    assert!(xml.contains("<vt:lpwstr>yes &amp; checked</vt:lpwstr>"));
  }

  #[test]
  fn adds_paragraph_text_before_section_properties() {
    let fixture = write_word_fixture();

    let bytes = add_paragraph_text(&fixture, "Added paragraph").expect("add paragraph");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(xml.contains("<w:t>Added paragraph</w:t></w:r></w:p><w:sectPr/>"));
  }

  #[test]
  fn inserts_table_before_section_properties() {
    let fixture = write_word_fixture();

    let bytes = insert_table(&fixture, &[&["A1", "B1"], &["A2", "B2"]]).expect("insert table");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(xml.contains("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t>"));
    assert!(xml.contains("<w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/>"));
  }

  #[test]
  fn changes_text_in_first_table_cell() {
    let fixture = write_word_fixture();

    let bytes =
      change_text_in_first_table_cell(&fixture, "Updated cell").expect("change table text");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(xml.contains("<w:t>Updated cell</w:t>"));
    assert!(!xml.contains("<w:t>Cell text</w:t>"));
  }

  #[test]
  fn sets_first_run_font() {
    let fixture = write_word_fixture();

    let bytes = set_first_run_font(&fixture, "Arial").expect("set font");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(
      xml.contains(r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Hello</w:t>"#)
    );
  }

  fn write_word_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-word-{}-{}.docx",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    let file = std::fs::File::create(&path).expect("create fixture");
    let mut zip = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();

    zip
      .start_file("[Content_Types].xml", options)
      .expect("content types");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"#,
    )
    .expect("write content types");

    zip.add_directory("_rels", options).expect("rels dir");
    zip
      .start_file("_rels/.rels", options)
      .expect("package rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#,
    )
    .expect("write package rels");

    zip
      .add_directory("docProps", options)
      .expect("docProps dir");
    zip
      .start_file("docProps/app.xml", options)
      .expect("app props");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>ooxmlsdk-doc</Application>
  <Pages>1</Pages>
  <Words>4</Words>
</Properties>"#,
      )
      .expect("write app props");

    zip.add_directory("word", options).expect("word dir");
    zip
      .add_directory("word/_rels", options)
      .expect("word rels dir");
    zip
      .start_file("word/_rels/document.xml.rels", options)
      .expect("document rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    )
    .expect("write document rels");

    zip
      .start_file("word/document.xml", options)
      .expect("document");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:p><w:r><w:t>from WordprocessingML</w:t></w:r></w:p>
    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"#,
      )
      .expect("write document");

    zip
      .start_file("word/comments.xml", options)
      .expect("comments");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Ada"><w:p><w:r><w:t>Review this paragraph</w:t></w:r></w:p></w:comment>
</w:comments>"#,
    )
    .expect("write comments");

    zip.start_file("word/styles.xml", options).expect("styles");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>"#,
      )
      .expect("write styles");

    zip.finish().expect("finish fixture");
    path
  }

  fn write_macro_enabled_word_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-word-macro-{}-{}.docm",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    let file = std::fs::File::create(&path).expect("create macro fixture");
    let mut zip = zip::ZipWriter::new(file);
    let options = zip::write::SimpleFileOptions::default();

    zip
      .start_file("[Content_Types].xml", options)
      .expect("content types");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/>
</Types>"#,
    )
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
      .start_file("word/_rels/document.xml.rels", options)
      .expect("document rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#,
    )
    .expect("write document rels");

    zip
      .start_file("word/document.xml", options)
      .expect("document");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Macro enabled</w:t></w:r></w:p></w:body>
</w:document>"#,
      )
      .expect("write document");

    zip
      .start_file("word/vbaProject.bin", options)
      .expect("vba project");
    zip.write_all(b"vba bytes").expect("write vba project");

    zip.finish().expect("finish macro fixture");
    path
  }

  fn write_bytes_fixture(extension: &str, bytes: Vec<u8>) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-word-bytes-{}-{}.{}",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed),
      extension
    ));
    std::fs::write(&path, bytes).expect("write bytes fixture");
    path
  }
}
