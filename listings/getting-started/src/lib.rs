// ANCHOR: full_example
use std::io::{Cursor, Write};

use ooxmlsdk::parts::wordprocessing_document::WordprocessingDocument;
use ooxmlsdk::schemas::schemas_openxmlformats_org_wordprocessingml_2006_main::{Body, Document};

pub fn create_empty_word_document() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut package = WordprocessingDocument::new(empty_open_xml_package()?)?;
  let main_part = package.add_main_document_part()?;

  main_part.set_root_element(
    &mut package,
    Document {
      body: Some(Box::new(Body::default())),
      ..Default::default()
    },
  )?;

  let mut buffer = Cursor::new(Vec::new());
  package.save(&mut buffer)?;
  Ok(buffer.into_inner())
}

fn empty_open_xml_package() -> Result<Cursor<Vec<u8>>, zip::result::ZipError> {
  let mut buffer = Cursor::new(Vec::new());
  {
    let mut zip = zip::ZipWriter::new(&mut buffer);
    let options = zip::write::SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
    )?;
    zip.finish()?;
  }
  buffer.set_position(0);
  Ok(buffer)
}
// ANCHOR_END: full_example

#[cfg(test)]
mod tests {
  use super::*;

  #[test]
  fn creates_reopenable_wordprocessing_document() {
    let bytes = create_empty_word_document().expect("document bytes");
    let mut reopened =
      WordprocessingDocument::new(Cursor::new(bytes)).expect("reopen generated docx");
    let main_part = reopened.main_document_part().expect("main document part");
    let document = main_part
      .root_element(&mut reopened)
      .expect("main document root");

    assert!(document.body.is_some());
  }
}
