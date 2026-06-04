// ANCHOR: open_word_read_only
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::header_part::HeaderPart;
use ooxmlsdk::parts::image_part::ImagePart;
use ooxmlsdk::parts::style_definitions_part::StyleDefinitionsPart;
use ooxmlsdk::parts::wordprocessing_comments_part::WordprocessingCommentsPart;
use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode, SdkPart, WordprocessingDocumentType};
use ooxmlsdk::validator::ValidationErrorInfo;

pub fn open_word_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();

  Ok(xml.matches("<w:p").count())
}
// ANCHOR_END: open_word_read_only

// ANCHOR: open_word_from_bytes
pub fn open_word_from_bytes(bytes: Vec<u8>) -> Result<usize, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new(Cursor::new(bytes))?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();

  Ok(xml.matches("<w:p").count())
}
// ANCHOR_END: open_word_from_bytes

// ANCHOR: validate_word_document
pub fn validate_word_document(
  path: &Path,
) -> Result<Vec<ValidationErrorInfo>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file(path)?;

  Ok(document.validate()?)
}
// ANCHOR_END: validate_word_document

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

// ANCHOR: replace_header
pub fn replace_header(
  path: &Path,
  header_text: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let header_part = if let Some(part) = main_part.header_parts(&document).next() {
    part
  } else {
    main_part.add_new_part_auto_id::<_, HeaderPart>(&mut document)?
  };
  let relationship_id = main_part
    .get_id_of_part(&document, &header_part)
    .expect("header relationship id")
    .to_string();

  header_part.set_data(&mut document, header_xml(header_text).into_bytes())?;
  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = set_header_reference(document_xml, &relationship_id)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: replace_header

// ANCHOR: remove_headers_and_footers
pub fn remove_headers_and_footers(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let header_parts: Vec<_> = main_part.header_parts(&document).collect();
  let footer_parts: Vec<_> = main_part.footer_parts(&document).collect();
  for header_part in header_parts {
    main_part.delete_part(&mut document, header_part)?;
  }
  for footer_part in footer_parts {
    main_part.delete_part(&mut document, footer_part)?;
  }

  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = remove_header_footer_references(document_xml);
  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: remove_headers_and_footers

// ANCHOR: create_paragraph_style
pub fn create_paragraph_style(
  path: &Path,
  style_id: &str,
  style_name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  add_style(path, style_id, style_name, "paragraph")
}
// ANCHOR_END: create_paragraph_style

// ANCHOR: create_character_style
pub fn create_character_style(
  path: &Path,
  style_id: &str,
  style_name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  add_style(path, style_id, style_name, "character")
}
// ANCHOR_END: create_character_style

// ANCHOR: apply_style_to_first_paragraph
pub fn apply_style_to_first_paragraph(
  path: &Path,
  style_id: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let Some(styles_part) = main_part.style_definitions_part(&document) else {
    return Err("document has no styles part".into());
  };
  let styles_xml = styles_part.data_as_str(&document)?.unwrap_or_default();
  if !styles_xml.contains(&format!(r#"w:styleId="{style_id}""#)) {
    return Err(format!("style {style_id} not found").into());
  }

  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = set_first_paragraph_style(document_xml, style_id)?;
  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: apply_style_to_first_paragraph

// ANCHOR: change_print_orientation
pub fn change_print_orientation(
  path: &Path,
  landscape: bool,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = set_section_orientation(document_xml, landscape)?;

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: change_print_orientation

// ANCHOR: replace_styles_part
pub fn replace_styles_part(
  path: &Path,
  styles_xml: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  if !styles_xml.contains("<w:styles") {
    return Err("styles XML must contain a w:styles root".into());
  }

  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let styles_part = if let Some(part) = main_part.style_definitions_part(&document) {
    part
  } else {
    main_part.add_new_part_auto_id::<_, StyleDefinitionsPart>(&mut document)?
  };

  styles_part.set_data(&mut document, styles_xml.as_bytes().to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: replace_styles_part

// ANCHOR: remove_hidden_text
pub fn remove_hidden_text(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = remove_runs_with_direct_vanish(document_xml);

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: remove_hidden_text

// ANCHOR: insert_comment
pub fn insert_comment(
  path: &Path,
  author: &str,
  comment_text: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let comments_part = if let Some(part) = main_part.wordprocessing_comments_part(&document) {
    part
  } else {
    main_part.add_new_part_auto_id::<_, WordprocessingCommentsPart>(&mut document)?
  };
  let comments_xml = comments_part
    .data_as_str(&document)?
    .map(str::to_string)
    .unwrap_or_else(empty_comments_xml);
  let comment_id = next_comment_id(&comments_xml);
  let updated_comments = append_comment(&comments_xml, comment_id, author, comment_text)?;
  comments_part.set_data(&mut document, updated_comments.into_bytes())?;

  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_document = add_comment_markers_to_first_paragraph(document_xml, comment_id)?;
  main_part.set_data(&mut document, updated_document.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_comment

// ANCHOR: delete_comments_by_author
pub fn delete_comments_by_author(
  path: &Path,
  author: Option<&str>,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let Some(comments_part) = main_part.wordprocessing_comments_part(&document) else {
    let mut buffer = Cursor::new(Vec::new());
    document.save(&mut buffer)?;
    return Ok(buffer.into_inner());
  };
  let comments_xml = comments_part.data_as_str(&document)?.unwrap_or_default();
  let deleted_ids = matching_comment_ids(comments_xml, author);
  let updated_comments = remove_comments_by_id(comments_xml, &deleted_ids);
  comments_part.set_data(&mut document, updated_comments.into_bytes())?;

  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_document = remove_comment_markers(document_xml, &deleted_ids);
  main_part.set_data(&mut document, updated_document.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: delete_comments_by_author

// ANCHOR: insert_picture
pub fn insert_picture(
  path: &Path,
  relationship_id: &str,
  content_type: &str,
  image_bytes: &[u8],
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let image_part =
    main_part.add_image_part_with_id(&mut document, content_type.to_string(), relationship_id)?;
  image_part.set_data(&mut document, image_bytes.to_vec())?;

  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let document_xml = ensure_relationship_namespace(document_xml);
  let picture = picture_paragraph_xml(relationship_id, 990_000, 792_000);
  let updated_xml = insert_before_section_properties(&document_xml, &picture)?;
  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_picture

// ANCHOR: replace_text_in_main_document
pub fn replace_text_in_main_document(
  path: &Path,
  search: &str,
  replacement: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = replace_text_values(document_xml, search, replacement);

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: replace_text_in_main_document

// ANCHOR: accept_common_revisions
pub fn accept_common_revisions(path: &Path) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let document_xml = main_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = accept_revision_markup(document_xml);

  main_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: accept_common_revisions

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
  let mut table = String::from("<w:tbl><w:tblPr></w:tblPr>");
  for row in rows {
    table.push_str("<w:tr>");
    for cell in *row {
      table.push_str("<w:tc><w:tcPr></w:tcPr>");
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

fn header_xml(text: &str) -> String {
  let text = escape_xml_text(text);
  format!(
    r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:hdr>"#
  )
}

fn add_style(
  path: &Path,
  style_id: &str,
  style_name: &str,
  style_type: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let styles_part = if let Some(part) = main_part.style_definitions_part(&document) {
    part
  } else {
    main_part.add_new_part_auto_id::<_, StyleDefinitionsPart>(&mut document)?
  };
  let styles_xml = styles_part
    .data_as_str(&document)?
    .map(str::to_string)
    .unwrap_or_else(|| {
      r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>"#
        .to_string()
    });
  let updated_xml = append_style(&styles_xml, style_id, style_name, style_type)?;

  styles_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}

fn append_style(
  styles_xml: &str,
  style_id: &str,
  style_name: &str,
  style_type: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  if styles_xml.contains(&format!(r#"w:styleId="{style_id}""#)) {
    return Ok(styles_xml.to_string());
  }
  let Some(insert_at) = styles_xml.find("</w:styles>") else {
    return Err("styles root not found".into());
  };
  let style_id = escape_xml_attr(style_id);
  let style_name = escape_xml_attr(style_name);
  let style_type = escape_xml_attr(style_type);
  let style_xml = format!(
    r#"<w:style w:type="{style_type}" w:styleId="{style_id}"><w:name w:val="{style_name}"/></w:style>"#
  );
  let mut updated = String::with_capacity(styles_xml.len() + style_xml.len());
  updated.push_str(&styles_xml[..insert_at]);
  updated.push_str(&style_xml);
  updated.push_str(&styles_xml[insert_at..]);
  Ok(updated)
}

fn ensure_relationship_namespace(document_xml: &str) -> String {
  if document_xml.contains("xmlns:r=") {
    return document_xml.to_string();
  }
  document_xml.replacen(
    "<w:document ",
    r#"<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    1,
  )
}

fn set_header_reference(
  document_xml: &str,
  relationship_id: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let document_xml = ensure_relationship_namespace(document_xml);
  let header_ref = format!(r#"<w:headerReference w:type="default" r:id="{relationship_id}"/>"#);
  if let Some((start, end)) = find_element_range(&document_xml, "<w:headerReference", None) {
    let mut updated = String::with_capacity(document_xml.len() + header_ref.len());
    updated.push_str(&document_xml[..start]);
    updated.push_str(&header_ref);
    updated.push_str(&document_xml[end..]);
    return Ok(updated);
  }
  let Some(sect_start) = document_xml.find("<w:sectPr") else {
    return Err("section properties not found".into());
  };
  let sect_open_end = document_xml[sect_start..]
    .find('>')
    .map(|index| sect_start + index + 1)
    .ok_or("section properties start not found")?;
  if document_xml[sect_open_end - 2..sect_open_end].starts_with('/') {
    let mut updated = String::with_capacity(document_xml.len() + header_ref.len() + 10);
    updated.push_str(&document_xml[..sect_open_end - 2]);
    updated.push('>');
    updated.push_str(&header_ref);
    updated.push_str("</w:sectPr>");
    updated.push_str(&document_xml[sect_open_end..]);
    return Ok(updated);
  }
  let mut updated = String::with_capacity(document_xml.len() + header_ref.len());
  updated.push_str(&document_xml[..sect_open_end]);
  updated.push_str(&header_ref);
  updated.push_str(&document_xml[sect_open_end..]);
  Ok(updated)
}

fn remove_header_footer_references(document_xml: &str) -> String {
  let without_headers = remove_all_elements(document_xml, "<w:headerReference");
  remove_all_elements(&without_headers, "<w:footerReference")
}

fn set_first_paragraph_style(
  document_xml: &str,
  style_id: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let paragraph_start = document_xml.find("<w:p>").ok_or("paragraph not found")?;
  let paragraph_open_end = paragraph_start + "<w:p>".len();
  let paragraph_end = document_xml[paragraph_open_end..]
    .find("</w:p>")
    .map(|index| paragraph_open_end + index)
    .ok_or("paragraph end not found")?;
  let paragraph_inner = &document_xml[paragraph_open_end..paragraph_end];
  let style_id = escape_xml_attr(style_id);
  let style_xml = format!(r#"<w:pStyle w:val="{style_id}"/>"#);

  let updated_inner = if let Some(ppr_start) = paragraph_inner.find("<w:pPr>") {
    let ppr_insert = ppr_start + "<w:pPr>".len();
    if let Some((style_start, style_end)) =
      find_element_range(paragraph_inner, "<w:pStyle", Some(ppr_insert))
    {
      let mut value = String::with_capacity(paragraph_inner.len() + style_xml.len());
      value.push_str(&paragraph_inner[..style_start]);
      value.push_str(&style_xml);
      value.push_str(&paragraph_inner[style_end..]);
      value
    } else {
      let mut value = String::with_capacity(paragraph_inner.len() + style_xml.len());
      value.push_str(&paragraph_inner[..ppr_insert]);
      value.push_str(&style_xml);
      value.push_str(&paragraph_inner[ppr_insert..]);
      value
    }
  } else {
    format!("<w:pPr>{style_xml}</w:pPr>{paragraph_inner}")
  };

  let mut updated = String::with_capacity(document_xml.len() + updated_inner.len());
  updated.push_str(&document_xml[..paragraph_open_end]);
  updated.push_str(&updated_inner);
  updated.push_str(&document_xml[paragraph_end..]);
  Ok(updated)
}

fn set_section_orientation(
  document_xml: &str,
  landscape: bool,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(sect_start) = document_xml.find("<w:sectPr") else {
    return Err("section properties not found".into());
  };
  let sect_open_end = document_xml[sect_start..]
    .find('>')
    .map(|index| sect_start + index + 1)
    .ok_or("section properties start not found")?;
  let (sect_end, was_self_closing) =
    if document_xml[sect_open_end - 2..sect_open_end].starts_with('/') {
      (sect_open_end, true)
    } else {
      (
        document_xml[sect_open_end..]
          .find("</w:sectPr>")
          .map(|index| sect_open_end + index + "</w:sectPr>".len())
          .ok_or("section properties end not found")?,
        false,
      )
    };
  let section_xml = if was_self_closing {
    let mut section = document_xml[sect_start..sect_open_end - 2].to_string();
    section.push('>');
    section.push_str("</w:sectPr>");
    section
  } else {
    document_xml[sect_start..sect_end].to_string()
  };
  let page_size = if landscape {
    r#"<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>"#
  } else {
    r#"<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>"#
  };
  let updated_section =
    if let Some((start, end)) = find_element_range(&section_xml, "<w:pgSz", None) {
      let mut section = String::with_capacity(section_xml.len() + page_size.len());
      section.push_str(&section_xml[..start]);
      section.push_str(page_size);
      section.push_str(&section_xml[end..]);
      section
    } else {
      let insert_at = section_xml.find('>').ok_or("section start not found")? + 1;
      let mut section = String::with_capacity(section_xml.len() + page_size.len());
      section.push_str(&section_xml[..insert_at]);
      section.push_str(page_size);
      section.push_str(&section_xml[insert_at..]);
      section
    };
  let mut updated = String::with_capacity(document_xml.len() + updated_section.len());
  updated.push_str(&document_xml[..sect_start]);
  updated.push_str(&updated_section);
  updated.push_str(&document_xml[sect_end..]);
  Ok(updated)
}

fn remove_all_elements(xml: &str, start_pattern: &str) -> String {
  let mut output = String::with_capacity(xml.len());
  let mut rest = xml;
  while let Some(start) = rest.find(start_pattern) {
    output.push_str(&rest[..start]);
    let element_rest = &rest[start..];
    let Some(open_end) = element_rest.find('>') else {
      output.push_str(element_rest);
      return output;
    };
    rest = &element_rest[open_end + 1..];
  }
  output.push_str(rest);
  output
}

fn find_element_range(
  xml: &str,
  start_pattern: &str,
  search_from: Option<usize>,
) -> Option<(usize, usize)> {
  let base = search_from.unwrap_or(0);
  let start = xml[base..].find(start_pattern)? + base;
  let open_end = xml[start..].find('>')? + start + 1;
  if xml[open_end - 2..open_end].starts_with('/') {
    return Some((start, open_end));
  }
  None
}

fn remove_runs_with_direct_vanish(document_xml: &str) -> String {
  let mut output = String::with_capacity(document_xml.len());
  let mut rest = document_xml;

  while let Some(run_start) = rest.find("<w:r") {
    output.push_str(&rest[..run_start]);
    let run_rest = &rest[run_start..];
    let Some(open_end) = run_rest.find('>') else {
      output.push_str(run_rest);
      return output;
    };
    let open_end = open_end + 1;
    if run_rest[open_end - 2..open_end].starts_with('/') {
      output.push_str(&run_rest[..open_end]);
      rest = &run_rest[open_end..];
      continue;
    }
    let Some(close_start) = run_rest[open_end..].find("</w:r>") else {
      output.push_str(run_rest);
      return output;
    };
    let run_end = open_end + close_start + "</w:r>".len();
    let run_xml = &run_rest[..run_end];
    if !has_direct_vanish(run_xml) {
      output.push_str(run_xml);
    }
    rest = &run_rest[run_end..];
  }

  output.push_str(rest);
  output
}

fn has_direct_vanish(run_xml: &str) -> bool {
  let Some(properties_start) = run_xml.find("<w:rPr") else {
    return false;
  };
  let Some(properties_open_end) = run_xml[properties_start..].find('>') else {
    return false;
  };
  let properties_open_end = properties_start + properties_open_end + 1;
  let Some(properties_close_start) = run_xml[properties_open_end..].find("</w:rPr>") else {
    return false;
  };
  let properties_xml = &run_xml[properties_open_end..properties_open_end + properties_close_start];
  let Some(vanish_start) = properties_xml.find("<w:vanish") else {
    return false;
  };
  let Some(vanish_open_end) = properties_xml[vanish_start..].find('>') else {
    return false;
  };
  let vanish_tag = &properties_xml[vanish_start..vanish_start + vanish_open_end + 1];
  !matches!(
    extract_attr(vanish_tag, "w:val"),
    Some("0" | "false" | "off")
  )
}

fn empty_comments_xml() -> String {
  r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:comments>"#
    .to_string()
}

fn next_comment_id(comments_xml: &str) -> u32 {
  let mut max_id = None;
  let mut rest = comments_xml;

  while let Some(start) = rest.find("<w:comment ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    if let Some(id) = extract_attr(&rest[..tag_end], "w:id").and_then(|value| value.parse().ok()) {
      max_id = Some(max_id.map_or(id, |current: u32| current.max(id)));
    }
    rest = &rest[tag_end + 1..];
  }

  max_id.map_or(0, |id| id + 1)
}

fn append_comment(
  comments_xml: &str,
  comment_id: u32,
  author: &str,
  comment_text: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(insert_at) = comments_xml.find("</w:comments>") else {
    return Err("comments root not found".into());
  };
  let author = escape_xml_attr(author);
  let comment_text = escape_xml_text(comment_text);
  let comment = format!(
    r#"<w:comment w:id="{comment_id}" w:author="{author}"><w:p><w:r><w:t>{comment_text}</w:t></w:r></w:p></w:comment>"#
  );
  let mut updated = String::with_capacity(comments_xml.len() + comment.len());
  updated.push_str(&comments_xml[..insert_at]);
  updated.push_str(&comment);
  updated.push_str(&comments_xml[insert_at..]);
  Ok(updated)
}

fn add_comment_markers_to_first_paragraph(
  document_xml: &str,
  comment_id: u32,
) -> Result<String, Box<dyn std::error::Error>> {
  let paragraph_start = document_xml.find("<w:p>").ok_or("paragraph not found")?;
  let paragraph_open_end = paragraph_start + "<w:p>".len();
  let paragraph_end = document_xml[paragraph_open_end..]
    .find("</w:p>")
    .map(|index| paragraph_open_end + index)
    .ok_or("paragraph end not found")?;
  let paragraph_inner = &document_xml[paragraph_open_end..paragraph_end];
  let run_start = paragraph_inner.find("<w:r").ok_or("run not found")?;
  let run_end = paragraph_inner[run_start..]
    .find("</w:r>")
    .map(|index| run_start + index + "</w:r>".len())
    .ok_or("run end not found")?;
  let markers_start = format!(r#"<w:commentRangeStart w:id="{comment_id}"/>"#);
  let markers_end = format!(
    r#"<w:commentRangeEnd w:id="{comment_id}"/><w:r><w:commentReference w:id="{comment_id}"/></w:r>"#
  );

  let mut updated_inner =
    String::with_capacity(paragraph_inner.len() + markers_start.len() + markers_end.len());
  updated_inner.push_str(&paragraph_inner[..run_start]);
  updated_inner.push_str(&markers_start);
  updated_inner.push_str(&paragraph_inner[run_start..run_end]);
  updated_inner.push_str(&markers_end);
  updated_inner.push_str(&paragraph_inner[run_end..]);

  let mut updated =
    String::with_capacity(document_xml.len() + markers_start.len() + markers_end.len());
  updated.push_str(&document_xml[..paragraph_open_end]);
  updated.push_str(&updated_inner);
  updated.push_str(&document_xml[paragraph_end..]);
  Ok(updated)
}

fn matching_comment_ids(comments_xml: &str, author: Option<&str>) -> Vec<String> {
  let mut ids = Vec::new();
  let mut rest = comments_xml;

  while let Some(start) = rest.find("<w:comment ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    let matches_author = author
      .map(|expected| extract_attr(tag, "w:author") == Some(expected))
      .unwrap_or(true);
    if matches_author && let Some(id) = extract_attr(tag, "w:id") {
      ids.push(id.to_string());
    }
    rest = &rest[tag_end + 1..];
  }

  ids
}

fn remove_comments_by_id(comments_xml: &str, ids: &[String]) -> String {
  let mut output = String::with_capacity(comments_xml.len());
  let mut rest = comments_xml;

  while let Some(start) = rest.find("<w:comment ") {
    output.push_str(&rest[..start]);
    let comment_rest = &rest[start..];
    let Some(tag_end) = comment_rest.find('>') else {
      output.push_str(comment_rest);
      return output;
    };
    let tag = &comment_rest[..tag_end];
    let Some(close_start) = comment_rest[tag_end + 1..].find("</w:comment>") else {
      output.push_str(comment_rest);
      return output;
    };
    let comment_end = tag_end + 1 + close_start + "</w:comment>".len();
    let should_delete =
      extract_attr(tag, "w:id").is_some_and(|id| ids.iter().any(|value| value == id));
    if !should_delete {
      output.push_str(&comment_rest[..comment_end]);
    }
    rest = &comment_rest[comment_end..];
  }

  output.push_str(rest);
  output
}

fn remove_comment_markers(document_xml: &str, ids: &[String]) -> String {
  let mut updated = document_xml.to_string();
  for id in ids {
    updated = remove_element_by_attr(&updated, "<w:commentRangeStart", "w:id", id);
    updated = remove_element_by_attr(&updated, "<w:commentRangeEnd", "w:id", id);
    updated = remove_element_by_attr(&updated, "<w:commentReference", "w:id", id);
  }
  updated
}

fn remove_element_by_attr(
  xml: &str,
  start_pattern: &str,
  attr_name: &str,
  attr_value: &str,
) -> String {
  let mut output = String::with_capacity(xml.len());
  let mut rest = xml;

  while let Some(start) = rest.find(start_pattern) {
    output.push_str(&rest[..start]);
    let element_rest = &rest[start..];
    let Some(open_end) = element_rest.find('>') else {
      output.push_str(element_rest);
      return output;
    };
    let tag = &element_rest[..open_end];
    if extract_attr(tag, attr_name) == Some(attr_value) {
      rest = &element_rest[open_end + 1..];
    } else {
      output.push_str(&element_rest[..open_end + 1]);
      rest = &element_rest[open_end + 1..];
    }
  }

  output.push_str(rest);
  output
}

fn picture_paragraph_xml(relationship_id: &str, width_emu: u64, height_emu: u64) -> String {
  let relationship_id = escape_xml_attr(relationship_id);
  format!(
    r#"<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{width_emu}" cy="{height_emu}"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
  )
}

fn replace_text_values(document_xml: &str, search: &str, replacement: &str) -> String {
  let replacement = escape_xml_text(replacement);
  let mut output = String::with_capacity(document_xml.len());
  let mut rest = document_xml;

  while let Some(start) = find_text_start(rest) {
    output.push_str(&rest[..start]);
    let text_rest = &rest[start..];
    let Some(open_end) = text_rest.find('>') else {
      output.push_str(text_rest);
      return output;
    };
    let open_end = open_end + 1;
    let Some(close_start) = text_rest[open_end..].find("</w:t>") else {
      output.push_str(text_rest);
      return output;
    };
    let close_start = open_end + close_start;
    let text = &text_rest[open_end..close_start];
    output.push_str(&text_rest[..open_end]);
    output.push_str(&text.replace(search, &replacement));
    output.push_str("</w:t>");
    rest = &text_rest[close_start + "</w:t>".len()..];
  }

  output.push_str(rest);
  output
}

fn accept_revision_markup(document_xml: &str) -> String {
  let mut accepted = document_xml.to_string();
  accepted = unwrap_all_elements(&accepted, "<w:ins ", "</w:ins>");
  accepted = unwrap_all_elements(&accepted, "<w:moveTo ", "</w:moveTo>");
  accepted = remove_all_paired_elements(&accepted, "<w:del ", "</w:del>");
  accepted = remove_all_paired_elements(&accepted, "<w:moveFrom ", "</w:moveFrom>");
  accepted = remove_all_paired_elements(&accepted, "<w:pPrChange ", "</w:pPrChange>");
  accepted = remove_all_paired_elements(&accepted, "<w:rPrChange ", "</w:rPrChange>");
  accepted = remove_all_elements(&accepted, "<w:moveFromRangeStart");
  accepted = remove_all_elements(&accepted, "<w:moveFromRangeEnd");
  accepted = remove_all_elements(&accepted, "<w:moveToRangeStart");
  remove_all_elements(&accepted, "<w:moveToRangeEnd")
}

fn unwrap_all_elements(xml: &str, start_pattern: &str, end_pattern: &str) -> String {
  let mut output = String::with_capacity(xml.len());
  let mut rest = xml;

  while let Some(start) = rest.find(start_pattern) {
    output.push_str(&rest[..start]);
    let element_rest = &rest[start..];
    let Some(open_end) = element_rest.find('>') else {
      output.push_str(element_rest);
      return output;
    };
    let open_end = open_end + 1;
    let Some(close_start) = element_rest[open_end..].find(end_pattern) else {
      output.push_str(element_rest);
      return output;
    };
    let close_start = open_end + close_start;
    output.push_str(&element_rest[open_end..close_start]);
    rest = &element_rest[close_start + end_pattern.len()..];
  }

  output.push_str(rest);
  output
}

fn remove_all_paired_elements(xml: &str, start_pattern: &str, end_pattern: &str) -> String {
  let mut output = String::with_capacity(xml.len());
  let mut rest = xml;

  while let Some(start) = rest.find(start_pattern) {
    output.push_str(&rest[..start]);
    let element_rest = &rest[start..];
    let Some(open_end) = element_rest.find('>') else {
      output.push_str(element_rest);
      return output;
    };
    if element_rest[open_end.saturating_sub(1)..=open_end].starts_with('/') {
      rest = &element_rest[open_end + 1..];
      continue;
    }
    let Some(close_start) = element_rest[open_end + 1..].find(end_pattern) else {
      output.push_str(element_rest);
      return output;
    };
    rest = &element_rest[open_end + 1 + close_start + end_pattern.len()..];
  }

  output.push_str(rest);
  output
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
  fn opens_word_from_bytes() {
    let bytes = std::fs::read(write_word_fixture()).expect("fixture bytes");

    let count = open_word_from_bytes(bytes).expect("open document from bytes");

    assert_eq!(count, 3);
  }

  #[test]
  fn validates_word_document() {
    let bytes = create_word_document("Valid text").expect("create document");
    let fixture = write_bytes_fixture("docx", bytes);

    let errors = validate_word_document(&fixture).expect("validate document");

    assert!(
      errors.is_empty(),
      "unexpected validation errors: {errors:?}"
    );
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

    assert!(
      xml
        .contains("<w:tbl><w:tblPr></w:tblPr><w:tr><w:tc><w:tcPr></w:tcPr><w:p><w:r><w:t>A1</w:t>")
    );
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

  #[test]
  fn replaces_default_header() {
    let fixture = write_word_fixture();

    let bytes = replace_header(&fixture, "Header Text").expect("replace header");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");
    let header_parts: Vec<_> = main_part.header_parts(&reopened).collect();
    let header_xml = header_parts[0]
      .data_as_str(&reopened)
      .expect("header xml")
      .expect("header data");

    assert_eq!(header_parts.len(), 1);
    assert!(document_xml.contains(r#"<w:headerReference w:type="default" r:id=""#));
    assert!(header_xml.contains("<w:t>Header Text</w:t>"));
  }

  #[test]
  fn removes_headers_and_footers() {
    let fixture = write_word_fixture();
    let bytes = replace_header(&fixture, "Header Text").expect("replace header");
    let path = write_bytes_fixture("docx", bytes);

    let bytes = remove_headers_and_footers(&path).expect("remove headers and footers");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert_eq!(main_part.header_parts(&reopened).count(), 0);
    assert_eq!(main_part.footer_parts(&reopened).count(), 0);
    assert!(!document_xml.contains("<w:headerReference"));
    assert!(!document_xml.contains("<w:footerReference"));
  }

  #[test]
  fn creates_paragraph_style() {
    let fixture = write_word_fixture();

    let bytes = create_paragraph_style(&fixture, "CodeBlock", "Code Block").expect("create style");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let styles_part = main_part
      .style_definitions_part(&reopened)
      .expect("styles part");
    let styles_xml = styles_part
      .data_as_str(&reopened)
      .expect("styles xml")
      .expect("styles data");

    assert!(styles_xml.contains(r#"<w:style w:type="paragraph" w:styleId="CodeBlock">"#));
    assert!(styles_xml.contains(r#"<w:name w:val="Code Block"/>"#));
  }

  #[test]
  fn creates_character_style() {
    let fixture = write_word_fixture();

    let bytes =
      create_character_style(&fixture, "EmphasisChar", "Emphasis Char").expect("create style");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let styles_part = main_part
      .style_definitions_part(&reopened)
      .expect("styles part");
    let styles_xml = styles_part
      .data_as_str(&reopened)
      .expect("styles xml")
      .expect("styles data");

    assert!(styles_xml.contains(r#"<w:style w:type="character" w:styleId="EmphasisChar">"#));
    assert!(styles_xml.contains(r#"<w:name w:val="Emphasis Char"/>"#));
  }

  #[test]
  fn applies_style_to_first_paragraph() {
    let fixture = write_word_fixture();

    let bytes = apply_style_to_first_paragraph(&fixture, "Heading1").expect("apply style");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(
      document_xml
        .contains(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello</w:t>"#)
    );
  }

  #[test]
  fn changes_print_orientation() {
    let fixture = write_word_fixture();

    let bytes = change_print_orientation(&fixture, true).expect("change orientation");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(document_xml.contains(r#"<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>"#));
  }

  #[test]
  fn replaces_styles_part() {
    let fixture = write_word_fixture();
    let replacement_styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/></w:style>
</w:styles>"#;

    let bytes = replace_styles_part(&fixture, replacement_styles).expect("replace styles");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let styles_part = main_part
      .style_definitions_part(&reopened)
      .expect("styles part");
    let styles_xml = styles_part
      .data_as_str(&reopened)
      .expect("styles xml")
      .expect("styles data");

    assert!(styles_xml.contains(r#"w:styleId="BodyText""#));
    assert!(!styles_xml.contains(r#"w:styleId="Heading1""#));
  }

  #[test]
  fn removes_hidden_text_runs() {
    let fixture = write_hidden_text_word_fixture();

    let bytes = remove_hidden_text(&fixture).expect("remove hidden text");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(document_xml.contains("<w:t>Visible</w:t>"));
    assert!(!document_xml.contains("Hidden"));
    assert!(!document_xml.contains("<w:vanish"));
  }

  #[test]
  fn inserts_comment_on_first_paragraph() {
    let fixture = write_word_fixture();

    let bytes = insert_comment(&fixture, "Grace", "Please review").expect("insert comment");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");
    let comments_part = main_part
      .wordprocessing_comments_part(&reopened)
      .expect("comments part");
    let comments_xml = comments_part
      .data_as_str(&reopened)
      .expect("comments xml")
      .expect("comments data");

    assert!(comments_xml.contains(r#"<w:comment w:id="1" w:author="Grace">"#));
    assert!(comments_xml.contains("<w:t>Please review</w:t>"));
    assert!(document_xml.contains(r#"<w:commentRangeStart w:id="1"/>"#));
    assert!(document_xml.contains(r#"<w:commentRangeEnd w:id="1"/>"#));
    assert!(document_xml.contains(r#"<w:commentReference w:id="1"/>"#));
  }

  #[test]
  fn deletes_comments_by_author() {
    let fixture = write_word_fixture();
    let bytes = insert_comment(&fixture, "Grace", "Please review").expect("insert comment");
    let path = write_bytes_fixture("docx", bytes);

    let bytes = delete_comments_by_author(&path, Some("Ada")).expect("delete comments");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");
    let comments_part = main_part
      .wordprocessing_comments_part(&reopened)
      .expect("comments part");
    let comments_xml = comments_part
      .data_as_str(&reopened)
      .expect("comments xml")
      .expect("comments data");

    assert!(!comments_xml.contains(r#"w:author="Ada""#));
    assert!(comments_xml.contains(r#"w:author="Grace""#));
    assert!(!document_xml.contains(r#"<w:commentReference w:id="0"/>"#));
    assert!(document_xml.contains(r#"<w:commentReference w:id="1"/>"#));
  }

  #[test]
  fn inserts_picture_markup_and_image_part() {
    let fixture = write_word_fixture();
    let image_bytes = b"\x89PNG\r\n\x1a\nimage bytes";

    let bytes =
      insert_picture(&fixture, "rIdPicture1", "image/png", image_bytes).expect("insert picture");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");
    let image_part = main_part
      .related_parts_of_type::<_, ImagePart>(&reopened)
      .find(|related| related.relationship_id() == "rIdPicture1")
      .map(|related| related.into_part())
      .expect("image part");

    assert_eq!(image_part.data(&reopened), Some(image_bytes.as_slice()));
    assert!(document_xml.contains(r#"r:embed="rIdPicture1""#));
    assert!(document_xml.contains("<wp:inline "));
    assert!(document_xml.contains("<pic:pic "));
  }

  #[test]
  fn replaces_text_in_main_document_text_nodes() {
    let fixture = write_word_fixture();

    let bytes =
      replace_text_in_main_document(&fixture, "Hello", "Hi & welcome").expect("replace text");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(document_xml.contains("<w:t>Hi &amp; welcome</w:t>"));
    assert!(!document_xml.contains("<w:t>Hello</w:t>"));
  }

  #[test]
  fn accepts_common_revision_markup() {
    let fixture = write_revision_word_fixture();

    let bytes = accept_common_revisions(&fixture).expect("accept revisions");

    let reopened = WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen document");
    let main_part = reopened.main_document_part().expect("main document part");
    let document_xml = main_part
      .data_as_str(&reopened)
      .expect("document xml")
      .expect("document data");

    assert!(document_xml.contains("<w:t>Kept insertion</w:t>"));
    assert!(document_xml.contains("<w:t>Moved here</w:t>"));
    assert!(!document_xml.contains("Deleted text"));
    assert!(!document_xml.contains("Moved away"));
    assert!(!document_xml.contains("<w:ins"));
    assert!(!document_xml.contains("<w:del"));
    assert!(!document_xml.contains("<w:moveFrom"));
    assert!(!document_xml.contains("<w:moveTo"));
    assert!(!document_xml.contains("<w:pPrChange"));
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
    <w:tbl><w:tblPr></w:tblPr><w:tr><w:tc><w:tcPr></w:tcPr><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
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

  fn write_hidden_text_word_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-word-hidden-{}-{}.docx",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    let file = std::fs::File::create(&path).expect("create hidden text fixture");
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
      .start_file("word/document.xml", options)
      .expect("document");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Visible</w:t></w:r>
      <w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#,
      )
      .expect("write document");

    zip.finish().expect("finish hidden text fixture");
    path
  }

  fn write_revision_word_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-word-revisions-{}-{}.docx",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    let file = std::fs::File::create(&path).expect("create revision fixture");
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
      .start_file("word/document.xml", options)
      .expect("document");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr>
      <w:ins w:id="2"><w:r><w:t>Kept insertion</w:t></w:r></w:ins>
      <w:del w:id="3"><w:r><w:delText>Deleted text</w:delText></w:r></w:del>
      <w:moveFromRangeStart w:id="4"/>
      <w:moveFrom w:id="5"><w:r><w:t>Moved away</w:t></w:r></w:moveFrom>
      <w:moveFromRangeEnd w:id="4"/>
      <w:moveToRangeStart w:id="6"/>
      <w:moveTo w:id="7"><w:r><w:t>Moved here</w:t></w:r></w:moveTo>
      <w:moveToRangeEnd w:id="6"/>
    </w:p>
  </w:body>
</w:document>"#,
      )
      .expect("write document");

    zip.finish().expect("finish revision fixture");
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
