// ANCHOR: open_presentation_read_only
use std::path::Path;

use ooxmlsdk::parts::presentation_document::PresentationDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode};

pub fn open_presentation_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;

  Ok(presentation_part.slide_parts(&document).count())
}
// ANCHOR_END: open_presentation_read_only

// ANCHOR: count_slides
pub fn count_slides(
  path: &Path,
  include_hidden: bool,
) -> Result<usize, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;

  if include_hidden {
    return Ok(presentation_part.slide_parts(&document).count());
  }

  let mut count = 0;
  for slide_part in presentation_part.slide_parts(&document) {
    let xml = slide_part.data_as_str(&document)?.unwrap_or_default();
    if !xml.contains(r#"show="0""#) && !xml.contains(r#"show="false""#) {
      count += 1;
    }
  }
  Ok(count)
}
// ANCHOR_END: count_slides

// ANCHOR: get_slide_text
pub fn get_slide_text(
  path: &Path,
  slide_index: usize,
) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Ok(Vec::new());
  };
  let xml = slide_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_drawing_text(xml))
}
// ANCHOR_END: get_slide_text

// ANCHOR: get_all_slide_text
pub fn get_all_slide_text(path: &Path) -> Result<Vec<Vec<String>>, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let mut slides = Vec::new();

  for slide_part in presentation_part.slide_parts(&document) {
    let xml = slide_part.data_as_str(&document)?.unwrap_or_default();
    slides.push(extract_drawing_text(xml));
  }

  Ok(slides)
}
// ANCHOR_END: get_all_slide_text

// ANCHOR: get_slide_titles
pub fn get_slide_titles(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let titles = get_all_slide_text(path)?
    .into_iter()
    .map(|slide_text| slide_text.into_iter().next().unwrap_or_default())
    .collect();

  Ok(titles)
}
// ANCHOR_END: get_slide_titles

// ANCHOR: get_external_hyperlinks
pub fn get_external_hyperlinks(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let mut links = Vec::new();

  for slide_part in presentation_part.slide_parts(&document) {
    let xml = slide_part.data_as_str(&document)?.unwrap_or_default();
    let hyperlink_ids = extract_hyperlink_relationship_ids(xml);

    for relationship in slide_part.hyperlink_relationships(&document) {
      if hyperlink_ids.iter().any(|id| id == relationship.id()) {
        links.push(relationship.target().to_string());
      }
    }
  }

  Ok(links)
}
// ANCHOR_END: get_external_hyperlinks

// ANCHOR: get_slide_layout_xml
pub fn get_slide_layout_xml(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let mut layouts = Vec::new();

  for slide_part in presentation_part.slide_parts(&document) {
    if let Some(layout_part) = slide_part.slide_layout_part(&document) {
      layouts.push(
        layout_part
          .data_as_str(&document)?
          .unwrap_or_default()
          .to_string(),
      );
    }
  }

  Ok(layouts)
}
// ANCHOR_END: get_slide_layout_xml

fn lazy_settings() -> OpenSettings {
  OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  }
}

fn extract_drawing_text(xml: &str) -> Vec<String> {
  let mut values = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<a:t>") {
    rest = &rest[start + "<a:t>".len()..];
    let Some(end) = rest.find("</a:t>") else {
      break;
    };
    values.push(decode_minimal_xml_text(&rest[..end]));
    rest = &rest[end + "</a:t>".len()..];
  }

  values
}

fn decode_minimal_xml_text(text: &str) -> String {
  text
    .replace("&lt;", "<")
    .replace("&gt;", ">")
    .replace("&quot;", "\"")
    .replace("&apos;", "'")
    .replace("&amp;", "&")
}

fn extract_hyperlink_relationship_ids(xml: &str) -> Vec<String> {
  let mut ids = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<a:hlink") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if let Some(id) = extract_attr(tag, "r:id") {
      ids.push(id.to_string());
    }
    rest = &rest[tag_end + 1..];
  }

  ids
}

fn extract_attr<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
  let pattern = format!(r#"{name}=""#);
  let start = tag.find(&pattern)? + pattern.len();
  let end = tag[start..].find('"')?;
  Some(&tag[start..start + end])
}

#[cfg(test)]
mod tests {
  use super::*;
  use std::io::Write;
  use std::sync::atomic::{AtomicUsize, Ordering};

  static FIXTURE_COUNTER: AtomicUsize = AtomicUsize::new(0);

  #[test]
  fn opens_presentation_read_only_and_counts_slide_parts() {
    let fixture = write_presentation_fixture();

    let count = open_presentation_read_only(&fixture).expect("open presentation");

    assert_eq!(count, 2);
  }

  #[test]
  fn counts_all_or_visible_slides() {
    let fixture = write_presentation_fixture();

    assert_eq!(count_slides(&fixture, true).expect("all slides"), 2);
    assert_eq!(count_slides(&fixture, false).expect("visible slides"), 1);
  }

  #[test]
  fn gets_text_from_slide() {
    let fixture = write_presentation_fixture();

    let text = get_slide_text(&fixture, 0).expect("slide text");

    assert_eq!(text, vec!["Intro", "Hello from slide 1", "Open intro link"]);
  }

  #[test]
  fn gets_text_from_all_slides() {
    let fixture = write_presentation_fixture();

    let text = get_all_slide_text(&fixture).expect("all slide text");

    assert_eq!(
      text,
      vec![
        vec![
          "Intro".to_string(),
          "Hello from slide 1".to_string(),
          "Open intro link".to_string()
        ],
        vec!["Hidden slide".to_string(), "Open hidden link".to_string()]
      ]
    );
  }

  #[test]
  fn gets_slide_titles() {
    let fixture = write_presentation_fixture();

    let titles = get_slide_titles(&fixture).expect("slide titles");

    assert_eq!(titles, vec!["Intro", "Hidden slide"]);
  }

  #[test]
  fn gets_external_hyperlinks() {
    let fixture = write_presentation_fixture();

    let links = get_external_hyperlinks(&fixture).expect("external hyperlinks");

    assert_eq!(
      links,
      vec![
        "https://example.com/intro".to_string(),
        "https://example.com/hidden".to_string()
      ]
    );
  }

  #[test]
  fn gets_slide_layout_xml() {
    let fixture = write_presentation_fixture();

    let layouts = get_slide_layout_xml(&fixture).expect("slide layout XML");

    assert_eq!(layouts.len(), 1);
    assert!(layouts[0].contains(r#"<p:sldLayout"#));
    assert!(layouts[0].contains(r#"type="title""#));
  }

  fn write_presentation_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-presentation-{}-{}.pptx",
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
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
    )
    .expect("write package rels");

    zip.add_directory("ppt", options).expect("ppt dir");
    zip
      .add_directory("ppt/_rels", options)
      .expect("ppt rels dir");
    zip
      .start_file("ppt/presentation.xml", options)
      .expect("presentation part");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#,
    )
    .expect("write presentation");

    zip
      .start_file("ppt/_rels/presentation.xml.rels", options)
      .expect("presentation rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>"#,
    )
    .expect("write presentation rels");

    zip
      .add_directory("ppt/slides", options)
      .expect("slides dir");
    zip
      .add_directory("ppt/slides/_rels", options)
      .expect("slide rels dir");
    write_slide(
      &mut zip,
      options,
      "ppt/slides/slide1.xml",
      None,
      &["Intro", "Hello from slide 1"],
      Some(("rLink1", "Open intro link")),
    );
    write_slide_rels(
      &mut zip,
      options,
      "ppt/slides/_rels/slide1.xml.rels",
      "rLink1",
      "https://example.com/intro",
      true,
    );
    write_slide(
      &mut zip,
      options,
      "ppt/slides/slide2.xml",
      Some("0"),
      &["Hidden slide"],
      Some(("rLink2", "Open hidden link")),
    );
    write_slide_rels(
      &mut zip,
      options,
      "ppt/slides/_rels/slide2.xml.rels",
      "rLink2",
      "https://example.com/hidden",
      false,
    );
    zip
      .add_directory("ppt/slideLayouts", options)
      .expect("slide layouts dir");
    zip
      .start_file("ppt/slideLayouts/slideLayout1.xml", options)
      .expect("slide layout part");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
  <p:cSld name="Title Slide"><p:spTree/></p:cSld>
</p:sldLayout>"#,
      )
      .expect("write slide layout");

    zip.finish().expect("finish fixture");
    path
  }

  fn write_slide<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: zip::write::SimpleFileOptions,
    path: &str,
    show: Option<&str>,
    text: &[&str],
    hyperlink: Option<(&str, &str)>,
  ) {
    zip.start_file(path, options).expect("slide part");
    let show_attr = show
      .map(|value| format!(r#" show="{value}""#))
      .unwrap_or_default();
    let text_xml = text
      .iter()
      .map(|value| {
        format!("<p:sp><p:txBody><a:p><a:r><a:t>{value}</a:t></a:r></a:p></p:txBody></p:sp>")
      })
      .collect::<String>();
    let hyperlink_xml = hyperlink
      .map(|(id, label)| {
        format!(
          r#"<p:sp><p:txBody><a:p><a:r><a:rPr><a:hlinkClick r:id="{id}"/></a:rPr><a:t>{label}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
      })
      .unwrap_or_default();
    let xml = format!(
      r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"{show_attr}>
  <p:cSld><p:spTree>{text_xml}{hyperlink_xml}</p:spTree></p:cSld>
</p:sld>"#
    );
    zip.write_all(xml.as_bytes()).expect("write slide");
  }

  fn write_slide_rels<W: std::io::Write + std::io::Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: zip::write::SimpleFileOptions,
    path: &str,
    id: &str,
    target: &str,
    include_layout: bool,
  ) {
    zip.start_file(path, options).expect("slide relationships");
    let layout_rel = if include_layout {
      r#"
  <Relationship Id="rLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>"#
    } else {
      ""
    };
    let xml = format!(
      r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{target}" TargetMode="External"/>
  <Relationship Id="rUnused" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/unused" TargetMode="External"/>{layout_rel}
</Relationships>"#
    );
    zip.write_all(xml.as_bytes()).expect("write slide rels");
  }
}
