// ANCHOR: open_presentation_read_only
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::presentation_document::PresentationDocument;
use ooxmlsdk::parts::presentation_part::PresentationPart;
use ooxmlsdk::parts::slide_part::SlidePart;
use ooxmlsdk::sdk::MediaDataPartType;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode, PresentationDocumentType};

pub fn open_presentation_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;

  Ok(presentation_part.slide_parts(&document).count())
}
// ANCHOR_END: open_presentation_read_only

// ANCHOR: create_presentation_document
pub fn create_presentation_document() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::create(PresentationDocumentType::Presentation);
  let presentation_part = document.add_new_part_auto_id::<PresentationPart>()?;
  let slide_part = presentation_part.add_new_part_auto_id::<_, SlidePart>(&mut document)?;
  let slide_relationship_id = presentation_part
    .get_id_of_part(&document, &slide_part)
    .expect("slide relationship id")
    .to_string();

  presentation_part.set_data(
    &mut document,
    format!(
      r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="{slide_relationship_id}"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    )
    .into_bytes(),
  )?;
  slide_part.set_data(
    &mut document,
    br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#.to_vec(),
  )?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: create_presentation_document

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

// ANCHOR: add_audio_media_references
pub fn add_audio_media_references(
  path: &Path,
  audio_bytes: &[u8],
) -> Result<(Vec<u8>, String, String), Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).next() else {
    return Err("presentation has no slide parts".into());
  };

  let media_part = document.create_media_data_part_by_type(MediaDataPartType::Wav)?;
  media_part.set_data(&mut document, audio_bytes.to_vec())?;
  let audio_relationship_id =
    slide_part.add_audio_reference_relationship(&mut document, &media_part)?;
  let media_relationship_id =
    slide_part.add_media_reference_relationship(&mut document, &media_part)?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok((
    buffer.into_inner(),
    audio_relationship_id,
    media_relationship_id,
  ))
}
// ANCHOR_END: add_audio_media_references

// ANCHOR: move_slide_to_position
pub fn move_slide_to_position(
  path: &Path,
  from_index: usize,
  to_index: usize,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let xml = presentation_part
    .data_as_str(&document)?
    .unwrap_or_default();
  let updated_xml = reorder_slide_ids(xml, from_index, to_index)?;

  presentation_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: move_slide_to_position

// ANCHOR: add_fade_transition
pub fn add_fade_transition(
  path: &Path,
  slide_index: usize,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Err(format!("slide index {slide_index} not found").into());
  };
  let xml = slide_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml =
    replace_or_insert_transition(xml, r#"<p:transition spd="fast"><p:fade/></p:transition>"#)?;

  slide_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_fade_transition

// ANCHOR: change_first_shape_fill_color
pub fn change_first_shape_fill_color(
  path: &Path,
  slide_index: usize,
  rgb_hex: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Err(format!("slide index {slide_index} not found").into());
  };
  let xml = slide_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = set_first_shape_solid_fill(xml, rgb_hex)?;

  slide_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: change_first_shape_fill_color

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

fn reorder_slide_ids(
  presentation_xml: &str,
  from_index: usize,
  to_index: usize,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(list_start) = presentation_xml.find("<p:sldIdLst>") else {
    return Err("presentation has no slide id list".into());
  };
  let list_content_start = list_start + "<p:sldIdLst>".len();
  let Some(list_end_offset) = presentation_xml[list_content_start..].find("</p:sldIdLst>") else {
    return Err("presentation slide id list is not closed".into());
  };
  let list_end = list_content_start + list_end_offset;
  let list_content = &presentation_xml[list_content_start..list_end];
  let mut slide_ids = collect_self_closing_elements(list_content, "<p:sldId");
  if from_index >= slide_ids.len() || to_index >= slide_ids.len() {
    return Err("slide index out of range".into());
  }

  let moved = slide_ids.remove(from_index);
  slide_ids.insert(to_index, moved);

  let mut updated = String::with_capacity(presentation_xml.len());
  updated.push_str(&presentation_xml[..list_content_start]);
  updated.push_str(&slide_ids.join(""));
  updated.push_str(&presentation_xml[list_end..]);
  Ok(updated)
}

fn replace_or_insert_transition(
  slide_xml: &str,
  transition_xml: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  if let Some((start, end)) = find_element_range(slide_xml, "<p:transition", "</p:transition>") {
    let mut updated = String::with_capacity(slide_xml.len() + transition_xml.len());
    updated.push_str(&slide_xml[..start]);
    updated.push_str(transition_xml);
    updated.push_str(&slide_xml[end..]);
    return Ok(updated);
  }

  let Some(common_slide_end) = slide_xml.find("</p:cSld>") else {
    return Err("slide has no common slide data".into());
  };
  let insert_at = common_slide_end + "</p:cSld>".len();
  let mut updated = String::with_capacity(slide_xml.len() + transition_xml.len());
  updated.push_str(&slide_xml[..insert_at]);
  updated.push_str(transition_xml);
  updated.push_str(&slide_xml[insert_at..]);
  Ok(updated)
}

fn set_first_shape_solid_fill(
  slide_xml: &str,
  rgb_hex: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  if rgb_hex.len() != 6 || !rgb_hex.chars().all(|ch| ch.is_ascii_hexdigit()) {
    return Err("rgb_hex must be six hexadecimal digits".into());
  }
  let shape_start = slide_xml.find("<p:sp>").ok_or("shape not found")?;
  let shape_end = slide_xml[shape_start..]
    .find("</p:sp>")
    .map(|index| shape_start + index + "</p:sp>".len())
    .ok_or("shape end not found")?;
  let shape_xml = &slide_xml[shape_start..shape_end];
  let solid_fill = format!(r#"<a:solidFill><a:srgbClr val="{rgb_hex}"/></a:solidFill>"#);

  let updated_shape =
    if let Some((sppr_start, sppr_end)) = find_element_range(shape_xml, "<p:spPr", "</p:spPr>") {
      let sppr_xml = &shape_xml[sppr_start..sppr_end];
      if let Some((fill_start, fill_end)) =
        find_element_range(sppr_xml, "<a:solidFill", "</a:solidFill>")
      {
        let mut updated_sppr = String::with_capacity(sppr_xml.len() + solid_fill.len());
        updated_sppr.push_str(&sppr_xml[..fill_start]);
        updated_sppr.push_str(&solid_fill);
        updated_sppr.push_str(&sppr_xml[fill_end..]);
        let mut updated_shape = String::with_capacity(shape_xml.len() + solid_fill.len());
        updated_shape.push_str(&shape_xml[..sppr_start]);
        updated_shape.push_str(&updated_sppr);
        updated_shape.push_str(&shape_xml[sppr_end..]);
        updated_shape
      } else {
        let insert_at = sppr_xml
          .find('>')
          .ok_or("shape properties start not found")?
          + 1;
        let mut updated_sppr = String::with_capacity(sppr_xml.len() + solid_fill.len());
        updated_sppr.push_str(&sppr_xml[..insert_at]);
        updated_sppr.push_str(&solid_fill);
        updated_sppr.push_str(&sppr_xml[insert_at..]);
        let mut updated_shape = String::with_capacity(shape_xml.len() + solid_fill.len());
        updated_shape.push_str(&shape_xml[..sppr_start]);
        updated_shape.push_str(&updated_sppr);
        updated_shape.push_str(&shape_xml[sppr_end..]);
        updated_shape
      }
    } else {
      let insert_at = shape_xml
        .find("<p:txBody")
        .ok_or("shape text body not found")?;
      let mut updated_shape = String::with_capacity(shape_xml.len() + solid_fill.len() + 15);
      updated_shape.push_str(&shape_xml[..insert_at]);
      updated_shape.push_str("<p:spPr>");
      updated_shape.push_str(&solid_fill);
      updated_shape.push_str("</p:spPr>");
      updated_shape.push_str(&shape_xml[insert_at..]);
      updated_shape
    };

  let mut updated = String::with_capacity(slide_xml.len() + updated_shape.len());
  updated.push_str(&slide_xml[..shape_start]);
  updated.push_str(&updated_shape);
  updated.push_str(&slide_xml[shape_end..]);
  Ok(updated)
}

fn collect_self_closing_elements(xml: &str, start_pattern: &str) -> Vec<String> {
  let mut elements = Vec::new();
  let mut rest = xml;
  while let Some(start) = rest.find(start_pattern) {
    rest = &rest[start..];
    let Some(end) = rest.find("/>") else {
      break;
    };
    elements.push(rest[..end + 2].to_string());
    rest = &rest[end + 2..];
  }
  elements
}

fn find_element_range(xml: &str, start_pattern: &str, end_pattern: &str) -> Option<(usize, usize)> {
  let start = xml.find(start_pattern)?;
  let open_end = xml[start..].find('>')? + start + 1;
  if xml[open_end - 2..open_end].starts_with('/') {
    return Some((start, open_end));
  }
  let end = xml[open_end..].find(end_pattern)? + open_end + end_pattern.len();
  Some((start, end))
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
  fn creates_presentation_document() {
    let bytes = create_presentation_document().expect("create presentation");
    let document =
      PresentationDocument::new(std::io::Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = document.presentation_part().expect("presentation part");

    assert_eq!(presentation_part.slide_parts(&document).count(), 1);
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

  #[test]
  fn adds_audio_media_references() {
    let fixture = write_presentation_fixture();
    let audio_bytes = b"RIFF....WAVEfmt audio bytes";

    let (bytes, audio_relationship_id, media_relationship_id) =
      add_audio_media_references(&fixture, audio_bytes).expect("add audio references");

    assert_ne!(audio_relationship_id, media_relationship_id);
    let reopened =
      PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation with media");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let slide_part = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let media_parts: Vec<_> = reopened.media_data_parts().collect();

    assert_eq!(media_parts.len(), 1);
    assert_eq!(media_parts[0].content_type(&reopened), Some("audio/wav"));
    assert_eq!(media_parts[0].data(&reopened), Some(audio_bytes.as_slice()));
    assert!(
      slide_part
        .data_part_reference_relationships(&reopened)
        .any(|relationship| relationship.id() == audio_relationship_id)
    );
    assert!(
      slide_part
        .data_part_reference_relationships(&reopened)
        .any(|relationship| relationship.id() == media_relationship_id)
    );
  }

  #[test]
  fn moves_slide_to_new_position() {
    let fixture = write_presentation_fixture();

    let bytes = move_slide_to_position(&fixture, 1, 0).expect("move slide");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let xml = presentation_part
      .data_as_str(&reopened)
      .expect("presentation xml")
      .expect("presentation data");
    let first = xml.find(r#"id="257""#).expect("second slide id");
    let second = xml.find(r#"id="256""#).expect("first slide id");

    assert!(first < second);
    assert_eq!(presentation_part.slide_parts(&reopened).count(), 2);
  }

  #[test]
  fn adds_fade_transition() {
    let fixture = write_presentation_fixture();

    let bytes = add_fade_transition(&fixture, 0).expect("add transition");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let slide_part = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let xml = slide_part
      .data_as_str(&reopened)
      .expect("slide xml")
      .expect("slide data");

    assert!(xml.contains(r#"<p:transition spd="fast"><p:fade/></p:transition>"#));
  }

  #[test]
  fn changes_first_shape_fill_color() {
    let fixture = write_presentation_fixture();

    let bytes = change_first_shape_fill_color(&fixture, 0, "FF0000").expect("change fill");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let slide_part = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let xml = slide_part
      .data_as_str(&reopened)
      .expect("slide xml")
      .expect("slide data");

    assert!(
      xml.contains(r#"<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>"#)
    );
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
  <p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
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
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{text_xml}{hyperlink_xml}</p:spTree></p:cSld>
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
