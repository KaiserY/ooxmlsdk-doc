// ANCHOR: open_presentation_read_only
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::comment_authors_part::CommentAuthorsPart;
use ooxmlsdk::parts::presentation_document::PresentationDocument;
use ooxmlsdk::parts::presentation_part::PresentationPart;
use ooxmlsdk::parts::slide_comments_part::SlideCommentsPart;
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
  let slide_parts: Vec<_> = presentation_part.slide_parts(&document).collect();
  for slide_part in slide_parts {
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

  let slide_parts: Vec<_> = presentation_part.slide_parts(&document).collect();
  for slide_part in slide_parts {
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

// ANCHOR: insert_new_slide
pub fn insert_new_slide(
  path: &Path,
  insertion_index: usize,
  title: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let slide_count = presentation_part.slide_parts(&document).count();
  if insertion_index > slide_count {
    return Err("insertion index out of range".into());
  }

  let layout_part = presentation_part
    .slide_parts(&document)
    .find_map(|slide| slide.slide_layout_part(&document));
  let slide_part = presentation_part.add_new_part_auto_id::<_, SlidePart>(&mut document)?;
  if let Some(layout_part) = layout_part {
    slide_part.create_relationship_to_part(&mut document, layout_part)?;
  }
  slide_part.set_data(&mut document, slide_xml(title).into_bytes())?;
  let relationship_id = presentation_part
    .get_id_of_part(&document, &slide_part)
    .expect("slide relationship id")
    .to_string();
  let presentation_xml = presentation_part
    .data_as_str(&document)?
    .unwrap_or_default();
  let updated_xml = insert_slide_id(presentation_xml, insertion_index, &relationship_id)?;

  presentation_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_new_slide

// ANCHOR: delete_slide
pub fn delete_slide(
  path: &Path,
  slide_index: usize,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let slides: Vec<_> = presentation_part.slide_parts(&document).collect();
  let Some(slide_part) = slides.get(slide_index).cloned() else {
    return Err("slide index out of range".into());
  };
  let presentation_xml = presentation_part
    .data_as_str(&document)?
    .unwrap_or_default();
  let updated_xml = remove_slide_id(presentation_xml, slide_index)?;

  presentation_part.delete_part(&mut document, slide_part)?;
  presentation_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: delete_slide

// ANCHOR: add_video_to_slide
pub fn add_video_to_slide(
  path: &Path,
  slide_index: usize,
  video_bytes: &[u8],
) -> Result<(Vec<u8>, String, String), Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Err("slide index out of range".into());
  };

  let media_part = document.create_media_data_part_by_type(MediaDataPartType::Mp4)?;
  media_part.set_data(&mut document, video_bytes.to_vec())?;
  let video_relationship_id =
    slide_part.add_video_reference_relationship(&mut document, &media_part)?;
  let media_relationship_id =
    slide_part.add_media_reference_relationship(&mut document, &media_part)?;
  let slide_xml = slide_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml =
    insert_video_picture(slide_xml, &video_relationship_id, &media_relationship_id)?;
  slide_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok((
    buffer.into_inner(),
    video_relationship_id,
    media_relationship_id,
  ))
}
// ANCHOR_END: add_video_to_slide

// ANCHOR: move_first_paragraph_between_presentations
pub fn move_first_paragraph_between_presentations(
  source_path: &Path,
  target_path: &Path,
) -> Result<(Vec<u8>, Vec<u8>), Box<dyn std::error::Error>> {
  let mut source = PresentationDocument::new_from_file_with_settings(source_path, lazy_settings())?;
  let mut target = PresentationDocument::new_from_file_with_settings(target_path, lazy_settings())?;
  let source_presentation_part = source.presentation_part()?;
  let target_presentation_part = target.presentation_part()?;
  let Some(source_slide_part) = source_presentation_part.slide_parts(&source).next() else {
    return Err("source presentation has no slides".into());
  };
  let Some(target_slide_part) = target_presentation_part.slide_parts(&target).next() else {
    return Err("target presentation has no slides".into());
  };

  let source_xml = source_slide_part.data_as_str(&source)?.unwrap_or_default();
  let (updated_source_xml, paragraph_xml) = remove_first_drawing_paragraph(source_xml)?;
  source_slide_part.set_data(&mut source, updated_source_xml.into_bytes())?;

  let target_xml = target_slide_part.data_as_str(&target)?.unwrap_or_default();
  let updated_target_xml = append_drawing_paragraph(target_xml, &paragraph_xml)?;
  target_slide_part.set_data(&mut target, updated_target_xml.into_bytes())?;

  let mut source_buffer = Cursor::new(Vec::new());
  source.save(&mut source_buffer)?;
  let mut target_buffer = Cursor::new(Vec::new());
  target.save(&mut target_buffer)?;
  Ok((source_buffer.into_inner(), target_buffer.into_inner()))
}
// ANCHOR_END: move_first_paragraph_between_presentations

// ANCHOR: add_comment_to_slide
pub fn add_comment_to_slide(
  path: &Path,
  slide_index: usize,
  author_name: &str,
  initials: &str,
  text: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let authors_part = if let Some(part) = presentation_part.comment_authors_part(&document) {
    part
  } else {
    presentation_part.add_new_part_auto_id::<_, CommentAuthorsPart>(&mut document)?
  };
  let authors_xml = authors_part
    .data_as_str(&document)?
    .map(str::to_string)
    .filter(|xml| !xml.trim().is_empty())
    .unwrap_or_else(empty_comment_authors_xml);
  let (updated_authors_xml, author_id, comment_index) =
    upsert_comment_author(&authors_xml, author_name, initials)?;
  authors_part.set_data(&mut document, updated_authors_xml.into_bytes())?;

  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Err("slide index out of range".into());
  };
  let comments_part = if let Some(part) = slide_part.slide_comments_part(&document) {
    part
  } else {
    slide_part.add_new_part_auto_id::<_, SlideCommentsPart>(&mut document)?
  };
  let comments_xml = comments_part
    .data_as_str(&document)?
    .map(str::to_string)
    .filter(|xml| !xml.trim().is_empty())
    .unwrap_or_else(empty_slide_comments_xml);
  let updated_comments_xml = append_slide_comment(&comments_xml, author_id, comment_index, text)?;
  comments_part.set_data(&mut document, updated_comments_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_comment_to_slide

// ANCHOR: delete_comments_by_author
pub fn delete_comments_by_author(
  path: &Path,
  author_name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(authors_part) = presentation_part.comment_authors_part(&document) else {
    let mut buffer = Cursor::new(Vec::new());
    document.save(&mut buffer)?;
    return Ok(buffer.into_inner());
  };
  let authors_xml = authors_part.data_as_str(&document)?.unwrap_or_default();
  let author_ids = comment_author_ids_by_name(authors_xml, author_name);
  let updated_authors_xml = remove_comment_authors_by_id(authors_xml, &author_ids);
  authors_part.set_data(&mut document, updated_authors_xml.into_bytes())?;

  let slide_parts: Vec<_> = presentation_part.slide_parts(&document).collect();
  for slide_part in slide_parts {
    let Some(comments_part) = slide_part.slide_comments_part(&document) else {
      continue;
    };
    let comments_xml = comments_part.data_as_str(&document)?.unwrap_or_default();
    let updated_comments_xml = remove_slide_comments_by_author_id(comments_xml, &author_ids);
    comments_part.set_data(&mut document, updated_comments_xml.into_bytes())?;
  }

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: delete_comments_by_author

// ANCHOR: add_basic_animation_timing
pub fn add_basic_animation_timing(
  path: &Path,
  slide_index: usize,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = PresentationDocument::new_from_file_with_settings(path, lazy_settings())?;
  let presentation_part = document.presentation_part()?;
  let Some(slide_part) = presentation_part.slide_parts(&document).nth(slide_index) else {
    return Err("slide index out of range".into());
  };
  let slide_xml = slide_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = replace_or_insert_timing(slide_xml, basic_timing_xml())?;
  slide_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: add_basic_animation_timing

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

fn slide_xml(title: &str) -> String {
  let title = escape_xml_text(title);
  format!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:txBody><a:p><a:r><a:t>{title}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>"#
  )
}

fn escape_xml_text(value: &str) -> String {
  value
    .replace('&', "&amp;")
    .replace('<', "&lt;")
    .replace('>', "&gt;")
}

fn insert_slide_id(
  presentation_xml: &str,
  insertion_index: usize,
  relationship_id: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some((content_start, list_end, slide_ids)) = slide_id_list(presentation_xml) else {
    return Err("presentation has no slide id list".into());
  };
  if insertion_index > slide_ids.len() {
    return Err("insertion index out of range".into());
  }
  let next_id = max_slide_id(&slide_ids) + 1;
  let mut updated_ids = slide_ids;
  updated_ids.insert(
    insertion_index,
    format!(r#"<p:sldId id="{next_id}" r:id="{relationship_id}"/>"#),
  );

  let mut updated = String::with_capacity(presentation_xml.len() + relationship_id.len() + 32);
  updated.push_str(&presentation_xml[..content_start]);
  updated.push_str(&updated_ids.join(""));
  updated.push_str(&presentation_xml[list_end..]);
  Ok(updated)
}

fn remove_slide_id(
  presentation_xml: &str,
  slide_index: usize,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some((content_start, list_end, mut slide_ids)) = slide_id_list(presentation_xml) else {
    return Err("presentation has no slide id list".into());
  };
  if slide_index >= slide_ids.len() {
    return Err("slide index out of range".into());
  }
  slide_ids.remove(slide_index);

  let mut updated = String::with_capacity(presentation_xml.len());
  updated.push_str(&presentation_xml[..content_start]);
  updated.push_str(&slide_ids.join(""));
  updated.push_str(&presentation_xml[list_end..]);
  Ok(updated)
}

fn slide_id_list(presentation_xml: &str) -> Option<(usize, usize, Vec<String>)> {
  let list_start = presentation_xml.find("<p:sldIdLst>")?;
  let content_start = list_start + "<p:sldIdLst>".len();
  let list_end = presentation_xml[content_start..].find("</p:sldIdLst>")? + content_start;
  let slide_ids =
    collect_self_closing_elements(&presentation_xml[content_start..list_end], "<p:sldId");
  Some((content_start, list_end, slide_ids))
}

fn max_slide_id(slide_ids: &[String]) -> u32 {
  slide_ids
    .iter()
    .filter_map(|slide_id| {
      let tag_end = slide_id.find('>')?;
      extract_attr(&slide_id[..tag_end], "id")?
        .parse::<u32>()
        .ok()
    })
    .max()
    .unwrap_or(255)
}

fn insert_video_picture(
  slide_xml: &str,
  video_relationship_id: &str,
  media_relationship_id: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(insert_at) = slide_xml.find("</p:spTree>") else {
    return Err("slide has no shape tree".into());
  };
  let video_relationship_id = escape_xml_attr(video_relationship_id);
  let media_relationship_id = escape_xml_attr(media_relationship_id);
  let picture = format!(
    r#"<p:pic><p:nvPicPr><p:cNvPr id="7" name="Video 1"><a:hlinkClick r:id="{media_relationship_id}" action="ppaction://media"/></p:cNvPr><p:cNvPicPr/><p:nvPr><a:videoFile r:link="{video_relationship_id}"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="{media_relationship_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2057400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
  );
  let mut updated = String::with_capacity(slide_xml.len() + picture.len());
  updated.push_str(&slide_xml[..insert_at]);
  updated.push_str(&picture);
  updated.push_str(&slide_xml[insert_at..]);
  Ok(updated)
}

fn escape_xml_attr(value: &str) -> String {
  escape_xml_text(value).replace('"', "&quot;")
}

fn remove_first_drawing_paragraph(
  slide_xml: &str,
) -> Result<(String, String), Box<dyn std::error::Error>> {
  let Some((start, end)) = first_drawing_paragraph_range(slide_xml) else {
    return Err("source slide has no drawing paragraph".into());
  };
  let paragraph = slide_xml[start..end].to_string();
  let mut updated = String::with_capacity(slide_xml.len());
  updated.push_str(&slide_xml[..start]);
  updated.push_str("<a:p/>");
  updated.push_str(&slide_xml[end..]);
  Ok((updated, paragraph))
}

fn append_drawing_paragraph(
  slide_xml: &str,
  paragraph_xml: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(body_start) = slide_xml.find("<p:txBody>") else {
    return Err("target slide has no text body".into());
  };
  let body_content_start = body_start + "<p:txBody>".len();
  let Some(body_end) = slide_xml[body_content_start..].find("</p:txBody>") else {
    return Err("target text body is not closed".into());
  };
  let insert_at = body_content_start + body_end;
  let mut updated = String::with_capacity(slide_xml.len() + paragraph_xml.len());
  updated.push_str(&slide_xml[..insert_at]);
  updated.push_str(paragraph_xml);
  updated.push_str(&slide_xml[insert_at..]);
  Ok(updated)
}

fn first_drawing_paragraph_range(slide_xml: &str) -> Option<(usize, usize)> {
  let start = slide_xml.find("<a:p")?;
  let open_end = slide_xml[start..].find('>')? + start + 1;
  if slide_xml[open_end - 2..open_end].starts_with('/') {
    return Some((start, open_end));
  }
  let end = slide_xml[open_end..].find("</a:p>")? + open_end + "</a:p>".len();
  Some((start, end))
}

fn empty_comment_authors_xml() -> String {
  r#"<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:cmAuthorLst>"#
    .to_string()
}

fn empty_slide_comments_xml() -> String {
  r#"<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:cmLst>"#
    .to_string()
}

fn upsert_comment_author(
  authors_xml: &str,
  author_name: &str,
  initials: &str,
) -> Result<(String, u32, u32), Box<dyn std::error::Error>> {
  if let Some((start, end, tag, author_id, last_index)) =
    find_comment_author(authors_xml, author_name, initials)
  {
    let next_index = last_index + 1;
    let updated_tag = set_or_add_attr(tag, "lastIdx", &next_index.to_string());
    let mut updated = String::with_capacity(authors_xml.len() + updated_tag.len());
    updated.push_str(&authors_xml[..start]);
    updated.push_str(&updated_tag);
    updated.push_str(&authors_xml[end..]);
    return Ok((updated, author_id, next_index));
  }

  let Some(insert_at) = authors_xml.find("</p:cmAuthorLst>") else {
    return Err("comment author list root not found".into());
  };
  let author_id = max_comment_author_id(authors_xml) + 1;
  let author_name = escape_xml_attr(author_name);
  let initials = escape_xml_attr(initials);
  let author_xml = format!(
    r#"<p:cmAuthor id="{author_id}" name="{author_name}" initials="{initials}" lastIdx="1" clrIdx="0"/>"#
  );
  let mut updated = String::with_capacity(authors_xml.len() + author_xml.len());
  updated.push_str(&authors_xml[..insert_at]);
  updated.push_str(&author_xml);
  updated.push_str(&authors_xml[insert_at..]);
  Ok((updated, author_id, 1))
}

fn append_slide_comment(
  comments_xml: &str,
  author_id: u32,
  comment_index: u32,
  text: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(insert_at) = comments_xml.find("</p:cmLst>") else {
    return Err("comment list root not found".into());
  };
  let text = escape_xml_text(text);
  let comment_xml = format!(
    r#"<p:cm authorId="{author_id}" dt="2026-05-26T00:00:00Z" idx="{comment_index}"><p:pos x="10" y="20"/><p:text>{text}</p:text></p:cm>"#
  );
  let mut updated = String::with_capacity(comments_xml.len() + comment_xml.len());
  updated.push_str(&comments_xml[..insert_at]);
  updated.push_str(&comment_xml);
  updated.push_str(&comments_xml[insert_at..]);
  Ok(updated)
}

fn find_comment_author<'a>(
  authors_xml: &'a str,
  author_name: &str,
  initials: &str,
) -> Option<(usize, usize, &'a str, u32, u32)> {
  let mut base = 0;
  let mut rest = authors_xml;
  while let Some(offset) = rest.find("<p:cmAuthor ") {
    let start = base + offset;
    let author_rest = &authors_xml[start..];
    let end = author_rest.find('>')? + start + 1;
    let tag = &authors_xml[start..end];
    if extract_attr(tag, "name") == Some(author_name)
      && extract_attr(tag, "initials") == Some(initials)
    {
      let author_id = extract_attr(tag, "id")?.parse().ok()?;
      let last_index = extract_attr(tag, "lastIdx")
        .and_then(|value| value.parse().ok())
        .unwrap_or(0);
      return Some((start, end, tag, author_id, last_index));
    }
    base = end;
    rest = &authors_xml[base..];
  }
  None
}

fn max_comment_author_id(authors_xml: &str) -> u32 {
  let mut max_id = 0;
  let mut rest = authors_xml;
  while let Some(start) = rest.find("<p:cmAuthor ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    if let Some(id) = extract_attr(&rest[..tag_end], "id").and_then(|value| value.parse().ok()) {
      max_id = max_id.max(id);
    }
    rest = &rest[tag_end + 1..];
  }
  max_id
}

fn comment_author_ids_by_name(authors_xml: &str, author_name: &str) -> Vec<String> {
  let mut ids = Vec::new();
  let mut rest = authors_xml;
  while let Some(start) = rest.find("<p:cmAuthor ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if extract_attr(tag, "name") == Some(author_name)
      && let Some(id) = extract_attr(tag, "id")
    {
      ids.push(id.to_string());
    }
    rest = &rest[tag_end + 1..];
  }
  ids
}

fn remove_comment_authors_by_id(authors_xml: &str, ids: &[String]) -> String {
  remove_elements_by_attr(authors_xml, "<p:cmAuthor ", "id", ids)
}

fn remove_slide_comments_by_author_id(comments_xml: &str, ids: &[String]) -> String {
  remove_paired_elements_by_attr(comments_xml, "<p:cm ", "</p:cm>", "authorId", ids)
}

fn remove_elements_by_attr(
  xml: &str,
  start_pattern: &str,
  attr_name: &str,
  ids: &[String],
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
    if extract_attr(tag, attr_name).is_some_and(|id| ids.iter().any(|value| value == id)) {
      rest = &element_rest[open_end + 1..];
    } else {
      output.push_str(&element_rest[..open_end + 1]);
      rest = &element_rest[open_end + 1..];
    }
  }
  output.push_str(rest);
  output
}

fn remove_paired_elements_by_attr(
  xml: &str,
  start_pattern: &str,
  end_pattern: &str,
  attr_name: &str,
  ids: &[String],
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
    let Some(close_start) = element_rest[open_end + 1..].find(end_pattern) else {
      output.push_str(element_rest);
      return output;
    };
    let element_end = open_end + 1 + close_start + end_pattern.len();
    let tag = &element_rest[..open_end];
    if extract_attr(tag, attr_name).is_some_and(|id| ids.iter().any(|value| value == id)) {
      rest = &element_rest[element_end..];
    } else {
      output.push_str(&element_rest[..element_end]);
      rest = &element_rest[element_end..];
    }
  }
  output.push_str(rest);
  output
}

fn replace_or_insert_timing(
  slide_xml: &str,
  timing_xml: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  if let Some((start, end)) = find_element_range(slide_xml, "<p:timing", "</p:timing>") {
    let mut updated = String::with_capacity(slide_xml.len() + timing_xml.len());
    updated.push_str(&slide_xml[..start]);
    updated.push_str(timing_xml);
    updated.push_str(&slide_xml[end..]);
    return Ok(updated);
  }

  let insert_at = slide_xml.find("</p:sld>").ok_or("slide root not found")?;
  let mut updated = String::with_capacity(slide_xml.len() + timing_xml.len());
  updated.push_str(&slide_xml[..insert_at]);
  updated.push_str(timing_xml);
  updated.push_str(&slide_xml[insert_at..]);
  Ok(updated)
}

fn basic_timing_xml() -> &'static str {
  r#"<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never"><p:childTnLst><p:par><p:cTn id="2" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="3" dur="500"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav></p:tavLst></p:anim></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"#
}

fn set_or_add_attr(tag: &str, name: &str, value: &str) -> String {
  if let Some(attr_start) = tag.find(&format!(r#"{name}=""#)) {
    let value_start = attr_start + name.len() + 2;
    if let Some(value_end) = tag[value_start..].find('"') {
      let value_end = value_start + value_end;
      let mut updated = String::with_capacity(tag.len() + value.len());
      updated.push_str(&tag[..value_start]);
      updated.push_str(value);
      updated.push_str(&tag[value_end..]);
      return updated;
    }
  }
  tag.replacen("/>", &format!(r#" {name}="{value}"/>"#), 1)
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

  #[test]
  fn inserts_new_slide_at_position() {
    let fixture = write_presentation_fixture();

    let bytes = insert_new_slide(&fixture, 1, "Inserted & Reviewed").expect("insert slide");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let presentation_xml = presentation_part
      .data_as_str(&reopened)
      .expect("presentation xml")
      .expect("presentation data");
    let titles = get_titles_from_document(&reopened).expect("titles");
    let first = presentation_xml
      .find(r#"id="256""#)
      .expect("first slide id");
    let inserted = presentation_xml
      .find(r#"id="258""#)
      .expect("inserted slide id");
    let last = presentation_xml
      .find(r#"id="257""#)
      .expect("second slide id");

    assert!(first < inserted);
    assert!(inserted < last);
    assert!(titles.contains(&"Inserted & Reviewed".to_string()));
    assert_eq!(presentation_part.slide_parts(&reopened).count(), 3);
  }

  #[test]
  fn deletes_slide_by_index() {
    let fixture = write_presentation_fixture();

    let bytes = delete_slide(&fixture, 1).expect("delete slide");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let titles = get_titles_from_document(&reopened).expect("titles");
    let presentation_xml = presentation_part
      .data_as_str(&reopened)
      .expect("presentation xml")
      .expect("presentation data");

    assert_eq!(titles, vec!["Intro"]);
    assert_eq!(presentation_part.slide_parts(&reopened).count(), 1);
    assert!(!presentation_xml.contains(r#"id="257""#));
    assert!(!presentation_xml.contains(r#"r:id="rId2""#));
  }

  #[test]
  fn adds_video_to_slide() {
    let fixture = write_presentation_fixture();
    let video_bytes = b"mp4 video bytes";

    let (bytes, video_relationship_id, media_relationship_id) =
      add_video_to_slide(&fixture, 0, video_bytes).expect("add video");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let slide_part = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let slide_xml = slide_part
      .data_as_str(&reopened)
      .expect("slide xml")
      .expect("slide data");
    let media_parts: Vec<_> = reopened.media_data_parts().collect();

    assert_eq!(media_parts.len(), 1);
    assert_eq!(media_parts[0].content_type(&reopened), Some("video/mp4"));
    assert_eq!(media_parts[0].data(&reopened), Some(video_bytes.as_slice()));
    assert!(slide_xml.contains(&format!(r#"a:videoFile r:link="{video_relationship_id}""#)));
    assert!(slide_xml.contains(&format!(r#"a:hlinkClick r:id="{media_relationship_id}""#)));
    assert!(
      slide_part
        .data_part_reference_relationships(&reopened)
        .any(|relationship| relationship.id() == video_relationship_id)
    );
    assert!(
      slide_part
        .data_part_reference_relationships(&reopened)
        .any(|relationship| relationship.id() == media_relationship_id)
    );
  }

  #[test]
  fn moves_first_paragraph_between_presentations() {
    let source_fixture = write_presentation_fixture();
    let target_fixture = write_presentation_fixture();

    let (source_bytes, target_bytes) =
      move_first_paragraph_between_presentations(&source_fixture, &target_fixture)
        .expect("move paragraph");

    let source = PresentationDocument::new(Cursor::new(source_bytes)).expect("reopen source");
    let target = PresentationDocument::new(Cursor::new(target_bytes)).expect("reopen target");
    let source_titles = get_titles_from_document(&source).expect("source titles");
    let target_part = target
      .presentation_part()
      .expect("target presentation part");
    let first_target_slide = target_part
      .slide_parts(&target)
      .next()
      .expect("target slide");
    let target_xml = first_target_slide
      .data_as_str(&target)
      .expect("target slide xml")
      .expect("target slide data");

    assert_eq!(source_titles[0], "Hello from slide 1");
    assert!(target_xml.matches("<a:t>Intro</a:t>").count() >= 2);
  }

  #[test]
  fn adds_comment_to_slide() {
    let fixture = write_presentation_fixture();

    let bytes =
      add_comment_to_slide(&fixture, 0, "Ada", "AL", "Review this slide").expect("add comment");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let authors_part = presentation_part
      .comment_authors_part(&reopened)
      .expect("comment authors part");
    let authors_xml = authors_part
      .data_as_str(&reopened)
      .expect("authors xml")
      .expect("authors data");
    let first_slide = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let comments_part = first_slide
      .slide_comments_part(&reopened)
      .expect("slide comments part");
    let comments_xml = comments_part
      .data_as_str(&reopened)
      .expect("comments xml")
      .expect("comments data");

    assert!(authors_xml.contains(r#"<p:cmAuthor id="1" name="Ada" initials="AL" lastIdx="1""#));
    assert!(comments_xml.contains(r#"<p:cm authorId="1""#));
    assert!(comments_xml.contains("<p:text>Review this slide</p:text>"));
  }

  #[test]
  fn deletes_comments_by_author() {
    let fixture = write_presentation_fixture();
    let bytes = add_comment_to_slide(&fixture, 0, "Ada", "AL", "Review this slide")
      .expect("add first comment");
    let path = write_bytes_fixture("pptx", bytes);
    let bytes = add_comment_to_slide(&path, 0, "Grace", "GH", "Keep this comment")
      .expect("add second comment");
    let path = write_bytes_fixture("pptx", bytes);

    let bytes = delete_comments_by_author(&path, "Ada").expect("delete comments");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let authors_part = presentation_part
      .comment_authors_part(&reopened)
      .expect("comment authors part");
    let authors_xml = authors_part
      .data_as_str(&reopened)
      .expect("authors xml")
      .expect("authors data");
    let first_slide = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let comments_part = first_slide
      .slide_comments_part(&reopened)
      .expect("slide comments part");
    let comments_xml = comments_part
      .data_as_str(&reopened)
      .expect("comments xml")
      .expect("comments data");

    assert!(!authors_xml.contains(r#"name="Ada""#));
    assert!(authors_xml.contains(r#"name="Grace""#));
    assert!(!comments_xml.contains("Review this slide"));
    assert!(comments_xml.contains("Keep this comment"));
  }

  #[test]
  fn adds_basic_animation_timing() {
    let fixture = write_presentation_fixture();

    let bytes = add_basic_animation_timing(&fixture, 0).expect("add animation timing");

    let reopened = PresentationDocument::new(Cursor::new(bytes)).expect("reopen presentation");
    let presentation_part = reopened.presentation_part().expect("presentation part");
    let first_slide = presentation_part
      .slide_parts(&reopened)
      .next()
      .expect("first slide");
    let slide_xml = first_slide
      .data_as_str(&reopened)
      .expect("slide xml")
      .expect("slide data");

    assert!(slide_xml.contains("<p:timing>"));
    assert!(slide_xml.contains(r#"<p:anim calcmode="lin" valueType="num">"#));
    assert!(slide_xml.contains(r#"<p:spTgt spid="2"/>"#));
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

  fn write_bytes_fixture(extension: &str, bytes: Vec<u8>) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-presentation-bytes-{}-{}.{}",
      std::process::id(),
      FIXTURE_COUNTER.fetch_add(1, Ordering::Relaxed),
      extension
    ));
    std::fs::write(&path, bytes).expect("write bytes fixture");
    path
  }

  fn get_titles_from_document(
    document: &PresentationDocument,
  ) -> Result<Vec<String>, Box<dyn std::error::Error>> {
    let presentation_part = document.presentation_part()?;
    let mut titles = Vec::new();
    for slide_part in presentation_part.slide_parts(document) {
      let xml = slide_part.data_as_str(document)?.unwrap_or_default();
      titles.push(
        extract_drawing_text(xml)
          .into_iter()
          .next()
          .unwrap_or_default(),
      );
    }
    Ok(titles)
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
