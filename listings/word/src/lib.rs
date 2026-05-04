// ANCHOR: open_word_read_only
use std::path::Path;

use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode};

pub fn open_word_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = WordprocessingDocument::new_from_file_with_settings(path, lazy_settings())?;
  let main_part = document.main_document_part()?;
  let xml = main_part.data_as_str(&document)?.unwrap_or_default();

  Ok(xml.matches("<w:p").count())
}
// ANCHOR_END: open_word_read_only

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

fn lazy_settings() -> OpenSettings {
  OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  }
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
}
