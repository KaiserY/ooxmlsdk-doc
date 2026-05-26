// ANCHOR: open_spreadsheet_read_only
use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::Path;

use ooxmlsdk::parts::ribbon_extensibility_part::RibbonExtensibilityPart;
use ooxmlsdk::parts::spreadsheet_document::SpreadsheetDocument;
use ooxmlsdk::parts::workbook_part::WorkbookPart;
use ooxmlsdk::parts::worksheet_part::WorksheetPart;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode, SdkPart, SpreadsheetDocumentType};

pub fn open_spreadsheet_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;

  Ok(workbook_part.worksheet_parts(&document).count())
}
// ANCHOR_END: open_spreadsheet_read_only

// ANCHOR: create_spreadsheet_document
pub fn create_spreadsheet_document() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::create(SpreadsheetDocumentType::Workbook);
  let workbook_part = document.add_new_part_auto_id::<WorkbookPart>()?;
  let worksheet_part = workbook_part.add_new_part_auto_id::<_, WorksheetPart>(&mut document)?;
  let worksheet_relationship_id = workbook_part
    .get_id_of_part(&document, &worksheet_part)
    .expect("worksheet relationship id")
    .to_string();

  workbook_part.set_data(
    &mut document,
    format!(
      r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="{worksheet_relationship_id}"/></sheets></workbook>"#
    )
    .into_bytes(),
  )?;
  worksheet_part.set_data(
    &mut document,
    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
  )?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: create_spreadsheet_document

// ANCHOR: list_worksheets
pub fn list_worksheets(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let workbook_xml = workbook_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_sheet_names(workbook_xml))
}
// ANCHOR_END: list_worksheets

// ANCHOR: get_worksheet_xml
pub fn get_worksheet_xml(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let mut worksheets = Vec::new();

  for worksheet_part in workbook_part.worksheet_parts(&document) {
    worksheets.push(
      worksheet_part
        .data_as_str(&document)?
        .unwrap_or_default()
        .to_string(),
    );
  }

  Ok(worksheets)
}
// ANCHOR_END: get_worksheet_xml

// ANCHOR: get_cell_values
pub fn get_cell_values(path: &Path) -> Result<Vec<(String, String)>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let shared_strings = workbook_part
    .shared_string_table_part(&document)
    .and_then(|part| {
      part
        .data_as_str(&document)
        .ok()
        .flatten()
        .map(extract_shared_strings)
    })
    .unwrap_or_default();
  let Some(first_sheet) = workbook_part.worksheet_parts(&document).next() else {
    return Ok(Vec::new());
  };
  let worksheet_xml = first_sheet.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_cell_values(worksheet_xml, &shared_strings))
}
// ANCHOR_END: get_cell_values

// ANCHOR: get_defined_names
pub fn get_defined_names(
  path: &Path,
) -> Result<BTreeMap<String, String>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let workbook_xml = workbook_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_defined_names(workbook_xml))
}
// ANCHOR_END: get_defined_names

// ANCHOR: get_hidden_rows_or_columns
pub fn get_hidden_rows_or_columns(
  path: &Path,
  sheet_name: &str,
  detect_rows: bool,
) -> Result<Vec<u32>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let workbook_xml = workbook_part.data_as_str(&document)?.unwrap_or_default();
  let Some(sheet_index) = extract_sheet_names(workbook_xml)
    .iter()
    .position(|name| name == sheet_name)
  else {
    return Ok(Vec::new());
  };
  let Some(worksheet_part) = workbook_part.worksheet_parts(&document).nth(sheet_index) else {
    return Ok(Vec::new());
  };
  let worksheet_xml = worksheet_part.data_as_str(&document)?.unwrap_or_default();

  Ok(if detect_rows {
    extract_hidden_rows(worksheet_xml)
  } else {
    extract_hidden_columns(worksheet_xml)
  })
}
// ANCHOR_END: get_hidden_rows_or_columns

// ANCHOR: get_hidden_worksheets
pub fn get_hidden_worksheets(path: &Path) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let workbook_xml = workbook_part.data_as_str(&document)?.unwrap_or_default();

  Ok(extract_hidden_sheet_names(workbook_xml))
}
// ANCHOR_END: get_hidden_worksheets

// ANCHOR: add_custom_ui_part
pub fn add_custom_ui_part(
  path: &Path,
  relationship_id: &str,
  custom_ui_xml: &[u8],
) -> Result<(Vec<u8>, String), Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let custom_ui_part = document.add_new_part::<RibbonExtensibilityPart>(relationship_id)?;

  custom_ui_part.set_data(&mut document, custom_ui_xml.to_vec())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok((buffer.into_inner(), relationship_id.to_string()))
}
// ANCHOR_END: add_custom_ui_part

// ANCHOR: list_worksheet_relationship_ids
pub fn list_worksheet_relationship_ids(
  path: &Path,
) -> Result<Vec<String>, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;

  Ok(
    workbook_part
      .related_parts_of_type::<_, WorksheetPart>(&document)
      .map(|related| related.relationship_id().to_string())
      .collect(),
  )
}
// ANCHOR_END: list_worksheet_relationship_ids

// ANCHOR: insert_text_into_cell
pub fn insert_text_into_cell(
  path: &Path,
  sheet_name: &str,
  cell_reference: &str,
  text: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let worksheet_part = worksheet_part_by_name(&document, &workbook_part, sheet_name)?;
  let worksheet_xml = worksheet_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = upsert_inline_string_cell(worksheet_xml, cell_reference, text)?;

  worksheet_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_text_into_cell

// ANCHOR: delete_text_from_cell
pub fn delete_text_from_cell(
  path: &Path,
  sheet_name: &str,
  cell_reference: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let worksheet_part = worksheet_part_by_name(&document, &workbook_part, sheet_name)?;
  let worksheet_xml = worksheet_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = clear_cell_text(worksheet_xml, cell_reference);

  worksheet_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: delete_text_from_cell

// ANCHOR: merge_adjacent_cells
pub fn merge_adjacent_cells(
  path: &Path,
  sheet_name: &str,
  first_cell: &str,
  second_cell: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let worksheet_part = worksheet_part_by_name(&document, &workbook_part, sheet_name)?;
  let worksheet_xml = worksheet_part.data_as_str(&document)?.unwrap_or_default();
  let updated_xml = add_merge_range(worksheet_xml, first_cell, second_cell)?;

  worksheet_part.set_data(&mut document, updated_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: merge_adjacent_cells

// ANCHOR: insert_new_worksheet
pub fn insert_new_worksheet(
  path: &Path,
  sheet_name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
  let mut document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;
  let worksheet_part = workbook_part.add_new_part_auto_id::<_, WorksheetPart>(&mut document)?;
  worksheet_part.set_data(
    &mut document,
    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
  )?;
  let relationship_id = workbook_part
    .get_id_of_part(&document, &worksheet_part)
    .expect("worksheet relationship id")
    .to_string();
  let workbook_xml = workbook_part.data_as_str(&document)?.unwrap_or_default();
  let updated_workbook_xml = append_sheet(workbook_xml, sheet_name, &relationship_id)?;

  workbook_part.set_data(&mut document, updated_workbook_xml.into_bytes())?;

  let mut buffer = Cursor::new(Vec::new());
  document.save(&mut buffer)?;
  Ok(buffer.into_inner())
}
// ANCHOR_END: insert_new_worksheet

fn lazy_settings() -> OpenSettings {
  OpenSettings {
    open_mode: PackageOpenMode::Lazy,
    ..Default::default()
  }
}

fn extract_sheet_names(xml: &str) -> Vec<String> {
  let mut names = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<sheet ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if let Some(name) = extract_attr(tag, "name") {
      names.push(decode_minimal_xml_text(name));
    }
    rest = &rest[tag_end + 1..];
  }

  names
}

fn extract_hidden_sheet_names(xml: &str) -> Vec<String> {
  let mut names = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<sheet ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if matches!(extract_attr(tag, "state"), Some("hidden" | "veryHidden"))
      && let Some(name) = extract_attr(tag, "name")
    {
      names.push(decode_minimal_xml_text(name));
    }
    rest = &rest[tag_end + 1..];
  }

  names
}

fn worksheet_part_by_name(
  document: &SpreadsheetDocument,
  workbook_part: &WorkbookPart,
  sheet_name: &str,
) -> Result<WorksheetPart, Box<dyn std::error::Error>> {
  let workbook_xml = workbook_part.data_as_str(document)?.unwrap_or_default();
  let Some(sheet_index) = extract_sheet_names(workbook_xml)
    .iter()
    .position(|name| name == sheet_name)
  else {
    return Err(format!("worksheet {sheet_name} not found").into());
  };
  workbook_part
    .worksheet_parts(document)
    .nth(sheet_index)
    .ok_or_else(|| format!("worksheet part for {sheet_name} not found").into())
}

fn upsert_inline_string_cell(
  worksheet_xml: &str,
  cell_reference: &str,
  text: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let row_index = row_index_from_cell_reference(cell_reference)?;
  let cell_xml = inline_string_cell_xml(cell_reference, text);

  if let Some((start, end)) = find_cell_range(worksheet_xml, cell_reference) {
    let mut updated = String::with_capacity(worksheet_xml.len() + cell_xml.len());
    updated.push_str(&worksheet_xml[..start]);
    updated.push_str(&cell_xml);
    updated.push_str(&worksheet_xml[end..]);
    return Ok(updated);
  }

  if let Some((row_start, row_open_end, row_end)) = find_row_range(worksheet_xml, row_index) {
    let mut updated = String::with_capacity(worksheet_xml.len() + cell_xml.len());
    updated.push_str(&worksheet_xml[..row_open_end]);
    updated.push_str(&cell_xml);
    updated.push_str(&worksheet_xml[row_open_end..row_end]);
    updated.push_str(&worksheet_xml[row_end..]);
    let _ = row_start;
    return Ok(updated);
  }

  let Some(sheet_data_end) = worksheet_xml.find("</sheetData>") else {
    return Err("worksheet has no sheetData element".into());
  };
  let row_xml = format!(r#"<row r="{row_index}">{cell_xml}</row>"#);
  let mut updated = String::with_capacity(worksheet_xml.len() + row_xml.len());
  updated.push_str(&worksheet_xml[..sheet_data_end]);
  updated.push_str(&row_xml);
  updated.push_str(&worksheet_xml[sheet_data_end..]);
  Ok(updated)
}

fn clear_cell_text(worksheet_xml: &str, cell_reference: &str) -> String {
  let Some((start, end)) = find_cell_range(worksheet_xml, cell_reference) else {
    return worksheet_xml.to_string();
  };
  let cell_xml = format!(r#"<c r="{cell_reference}"/>"#);
  let mut updated = String::with_capacity(worksheet_xml.len());
  updated.push_str(&worksheet_xml[..start]);
  updated.push_str(&cell_xml);
  updated.push_str(&worksheet_xml[end..]);
  updated
}

fn add_merge_range(
  worksheet_xml: &str,
  first_cell: &str,
  second_cell: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let merge_reference = merge_reference(first_cell, second_cell)?;
  if worksheet_xml.contains(&format!(r#"ref="{merge_reference}""#)) {
    return Ok(worksheet_xml.to_string());
  }

  if let Some((start, open_end, end)) = find_merge_cells_range(worksheet_xml) {
    let merge_xml = format!(r#"<mergeCell ref="{merge_reference}"/>"#);
    let existing = &worksheet_xml[open_end..end];
    let count = existing.matches("<mergeCell ").count() + 1;
    let mut opening = worksheet_xml[start..open_end].to_string();
    opening = set_or_add_attr(&opening, "count", &count.to_string());

    let mut updated = String::with_capacity(worksheet_xml.len() + merge_xml.len());
    updated.push_str(&worksheet_xml[..start]);
    updated.push_str(&opening);
    updated.push_str(existing);
    updated.push_str(&merge_xml);
    updated.push_str(&worksheet_xml[end..]);
    return Ok(updated);
  }

  let insert_at = worksheet_xml
    .find("</sheetData>")
    .map(|index| index + "</sheetData>".len())
    .ok_or("worksheet has no sheetData element")?;
  let merge_xml =
    format!(r#"<mergeCells count="1"><mergeCell ref="{merge_reference}"/></mergeCells>"#);
  let mut updated = String::with_capacity(worksheet_xml.len() + merge_xml.len());
  updated.push_str(&worksheet_xml[..insert_at]);
  updated.push_str(&merge_xml);
  updated.push_str(&worksheet_xml[insert_at..]);
  Ok(updated)
}

fn append_sheet(
  workbook_xml: &str,
  sheet_name: &str,
  relationship_id: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let Some(sheets_end) = workbook_xml.find("</sheets>") else {
    return Err("workbook has no sheets element".into());
  };
  let next_sheet_id = max_sheet_id(workbook_xml) + 1;
  let sheet_name = escape_xml_text(sheet_name);
  let sheet_xml =
    format!(r#"<sheet name="{sheet_name}" sheetId="{next_sheet_id}" r:id="{relationship_id}"/>"#);
  let mut updated = String::with_capacity(workbook_xml.len() + sheet_xml.len());
  updated.push_str(&workbook_xml[..sheets_end]);
  updated.push_str(&sheet_xml);
  updated.push_str(&workbook_xml[sheets_end..]);
  Ok(updated)
}

fn extract_shared_strings(xml: &str) -> Vec<String> {
  let mut values = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<si>") {
    rest = &rest[start + "<si>".len()..];
    let Some(end) = rest.find("</si>") else {
      break;
    };
    values.push(extract_text_values(&rest[..end]).join(""));
    rest = &rest[end + "</si>".len()..];
  }

  values
}

fn extract_cell_values(xml: &str, shared_strings: &[String]) -> Vec<(String, String)> {
  let mut cells = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<c ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let cell_tag = &rest[..tag_end];
    let reference = extract_attr(cell_tag, "r").unwrap_or_default().to_string();
    let data_type = extract_attr(cell_tag, "t");
    let Some(cell_end) = rest.find("</c>") else {
      rest = &rest[tag_end + 1..];
      continue;
    };
    let cell_xml = &rest[tag_end + 1..cell_end];
    if let Some(raw_value) = extract_element_text(cell_xml, "v") {
      let value = if data_type == Some("s") {
        raw_value
          .parse::<usize>()
          .ok()
          .and_then(|index| shared_strings.get(index))
          .cloned()
          .unwrap_or_default()
      } else {
        decode_minimal_xml_text(raw_value)
      };
      cells.push((reference, value));
    }
    rest = &rest[cell_end + "</c>".len()..];
  }

  cells
}

fn extract_defined_names(xml: &str) -> BTreeMap<String, String> {
  let mut names = BTreeMap::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<definedName ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    let Some(name) = extract_attr(tag, "name") else {
      rest = &rest[tag_end + 1..];
      continue;
    };
    let Some(end) = rest.find("</definedName>") else {
      break;
    };
    names.insert(
      decode_minimal_xml_text(name),
      decode_minimal_xml_text(&rest[tag_end + 1..end]),
    );
    rest = &rest[end + "</definedName>".len()..];
  }

  names
}

fn extract_hidden_rows(xml: &str) -> Vec<u32> {
  let mut rows = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<row ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if is_hidden(tag)
      && let Some(index) = extract_attr(tag, "r").and_then(|value| value.parse::<u32>().ok())
    {
      rows.push(index);
    }
    rest = &rest[tag_end + 1..];
  }

  rows
}

fn extract_hidden_columns(xml: &str) -> Vec<u32> {
  let mut columns = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<col ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    let tag = &rest[..tag_end];
    if is_hidden(tag) {
      let min = extract_attr(tag, "min").and_then(|value| value.parse::<u32>().ok());
      let max = extract_attr(tag, "max").and_then(|value| value.parse::<u32>().ok());
      if let (Some(min), Some(max)) = (min, max) {
        columns.extend(min..=max);
      }
    }
    rest = &rest[tag_end + 1..];
  }

  columns
}

fn is_hidden(tag: &str) -> bool {
  matches!(extract_attr(tag, "hidden"), Some("1" | "true"))
}

fn find_cell_range(xml: &str, cell_reference: &str) -> Option<(usize, usize)> {
  let pattern = format!(r#"<c r="{cell_reference}""#);
  let start = xml.find(&pattern)?;
  let open_end = xml[start..].find('>')? + start + 1;
  if xml[open_end - 2..open_end].starts_with("/") {
    return Some((start, open_end));
  }
  let end = xml[open_end..].find("</c>")? + open_end + "</c>".len();
  Some((start, end))
}

fn find_row_range(xml: &str, row_index: u32) -> Option<(usize, usize, usize)> {
  let pattern = format!(r#"<row r="{row_index}""#);
  let start = xml.find(&pattern)?;
  let open_end = xml[start..].find('>')? + start + 1;
  if xml[open_end - 2..open_end].starts_with("/") {
    return Some((start, open_end - 2, open_end));
  }
  let end = xml[open_end..].find("</row>")? + open_end;
  Some((start, open_end, end))
}

fn find_merge_cells_range(xml: &str) -> Option<(usize, usize, usize)> {
  let start = xml.find("<mergeCells")?;
  let open_end = xml[start..].find('>')? + start + 1;
  let end = xml[open_end..].find("</mergeCells>")? + open_end;
  Some((start, open_end, end))
}

fn inline_string_cell_xml(cell_reference: &str, text: &str) -> String {
  let text = escape_xml_text(text);
  format!(r#"<c r="{cell_reference}" t="inlineStr"><is><t>{text}</t></is></c>"#)
}

fn row_index_from_cell_reference(cell_reference: &str) -> Result<u32, Box<dyn std::error::Error>> {
  let digits: String = cell_reference
    .chars()
    .skip_while(|ch| ch.is_ascii_alphabetic())
    .collect();
  let row = digits.parse::<u32>()?;
  if row == 0 {
    return Err("cell row index must be greater than zero".into());
  }
  Ok(row)
}

fn column_index_from_cell_reference(
  cell_reference: &str,
) -> Result<u32, Box<dyn std::error::Error>> {
  let mut column = 0_u32;
  for ch in cell_reference
    .chars()
    .take_while(|ch| ch.is_ascii_alphabetic())
  {
    column = column * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
  }
  if column == 0 {
    return Err("cell reference has no column".into());
  }
  Ok(column)
}

fn merge_reference(
  first_cell: &str,
  second_cell: &str,
) -> Result<String, Box<dyn std::error::Error>> {
  let first_row = row_index_from_cell_reference(first_cell)?;
  let second_row = row_index_from_cell_reference(second_cell)?;
  let first_column = column_index_from_cell_reference(first_cell)?;
  let second_column = column_index_from_cell_reference(second_cell)?;

  let adjacent = (first_row == second_row && first_column.abs_diff(second_column) == 1)
    || (first_column == second_column && first_row.abs_diff(second_row) == 1);
  if !adjacent {
    return Err(format!("{first_cell} and {second_cell} are not adjacent").into());
  }

  let (start, end) = if (first_row, first_column) <= (second_row, second_column) {
    (first_cell, second_cell)
  } else {
    (second_cell, first_cell)
  };
  Ok(format!("{start}:{end}"))
}

fn max_sheet_id(workbook_xml: &str) -> u32 {
  let mut max_id = 0_u32;
  let mut rest = workbook_xml;
  while let Some(start) = rest.find("<sheet ") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    if let Some(sheet_id) =
      extract_attr(&rest[..tag_end], "sheetId").and_then(|value| value.parse::<u32>().ok())
    {
      max_id = max_id.max(sheet_id);
    }
    rest = &rest[tag_end + 1..];
  }
  max_id
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
  tag.replacen('>', &format!(r#" count="{value}">"#), 1)
}

fn extract_text_values(xml: &str) -> Vec<String> {
  let mut values = Vec::new();
  let mut rest = xml;

  while let Some(start) = rest.find("<t") {
    rest = &rest[start..];
    let Some(tag_end) = rest.find('>') else {
      break;
    };
    rest = &rest[tag_end + 1..];
    let Some(end) = rest.find("</t>") else {
      break;
    };
    values.push(decode_minimal_xml_text(&rest[..end]));
    rest = &rest[end + "</t>".len()..];
  }

  values
}

fn extract_element_text<'a>(xml: &'a str, name: &str) -> Option<&'a str> {
  let open = format!("<{name}>");
  let close = format!("</{name}>");
  let start = xml.find(&open)? + open.len();
  let end = xml[start..].find(&close)?;
  Some(&xml[start..start + end])
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

fn escape_xml_text(value: &str) -> String {
  value
    .replace('&', "&amp;")
    .replace('<', "&lt;")
    .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {
  use super::*;
  use std::io::Write;
  use std::sync::atomic::{AtomicUsize, Ordering};

  static FIXTURE_COUNTER: AtomicUsize = AtomicUsize::new(0);

  #[test]
  fn opens_spreadsheet_read_only_and_counts_worksheets() {
    let fixture = write_spreadsheet_fixture();

    let count = open_spreadsheet_read_only(&fixture).expect("open spreadsheet");

    assert_eq!(count, 2);
  }

  #[test]
  fn creates_spreadsheet_document() {
    let bytes = create_spreadsheet_document().expect("create spreadsheet");
    let document =
      SpreadsheetDocument::new(std::io::Cursor::new(bytes)).expect("reopen spreadsheet");
    let workbook_part = document.workbook_part().expect("workbook part");

    assert_eq!(workbook_part.worksheet_parts(&document).count(), 1);
  }

  #[test]
  fn lists_worksheets() {
    let fixture = write_spreadsheet_fixture();

    let sheets = list_worksheets(&fixture).expect("worksheet names");

    assert_eq!(sheets, vec!["Summary", "Hidden Data"]);
  }

  #[test]
  fn gets_worksheet_xml() {
    let fixture = write_spreadsheet_fixture();

    let worksheets = get_worksheet_xml(&fixture).expect("worksheet XML");

    assert_eq!(worksheets.len(), 2);
    assert!(worksheets[0].contains(r#"<worksheet"#));
    assert!(worksheets[1].contains(r#"state="hidden""#));
  }

  #[test]
  fn gets_cell_values() {
    let fixture = write_spreadsheet_fixture();

    let values = get_cell_values(&fixture).expect("cell values");

    assert_eq!(
      values,
      vec![
        ("A1".to_string(), "Region".to_string()),
        ("B1".to_string(), "Sales".to_string()),
        ("A2".to_string(), "North".to_string()),
        ("B2".to_string(), "42".to_string())
      ]
    );
  }

  #[test]
  fn gets_defined_names() {
    let fixture = write_spreadsheet_fixture();

    let names = get_defined_names(&fixture).expect("defined names");

    assert_eq!(
      names.get("SalesRange").map(String::as_str),
      Some("Summary!$B$2:$B$2")
    );
  }

  #[test]
  fn gets_hidden_rows_or_columns() {
    let fixture = write_spreadsheet_fixture();

    let rows = get_hidden_rows_or_columns(&fixture, "Summary", true).expect("hidden rows");
    let columns = get_hidden_rows_or_columns(&fixture, "Summary", false).expect("hidden columns");

    assert_eq!(rows, vec![3]);
    assert_eq!(columns, vec![3, 4]);
  }

  #[test]
  fn gets_hidden_worksheets() {
    let fixture = write_spreadsheet_fixture();

    let sheets = get_hidden_worksheets(&fixture).expect("hidden worksheets");

    assert_eq!(sheets, vec!["Hidden Data"]);
  }

  #[test]
  fn adds_custom_ui_part() {
    let fixture = write_spreadsheet_fixture();
    let custom_ui = br#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>"#;

    let (bytes, relationship_id) =
      add_custom_ui_part(&fixture, "rIdCustomUi1", custom_ui).expect("add custom UI");

    assert_eq!(relationship_id, "rIdCustomUi1");
    let reopened =
      SpreadsheetDocument::new(Cursor::new(bytes)).expect("reopen spreadsheet with custom UI");
    let custom_part = reopened
      .related_parts_of_type::<RibbonExtensibilityPart>()
      .find(|related| related.relationship_id() == "rIdCustomUi1")
      .map(|related| related.into_part())
      .expect("custom UI part");

    assert_eq!(custom_part.data(&reopened), Some(custom_ui.as_slice()));
  }

  #[test]
  fn lists_worksheet_relationship_ids() {
    let fixture = write_spreadsheet_fixture();

    let ids = list_worksheet_relationship_ids(&fixture).expect("worksheet relationship ids");

    assert_eq!(ids, vec!["rId1", "rId2"]);
  }

  #[test]
  fn inserts_text_into_cell() {
    let fixture = write_spreadsheet_fixture();

    let bytes =
      insert_text_into_cell(&fixture, "Summary", "C4", "New & Value").expect("insert text");

    let reopened = SpreadsheetDocument::new(Cursor::new(bytes)).expect("reopen spreadsheet");
    let workbook_part = reopened.workbook_part().expect("workbook part");
    let first_sheet = workbook_part
      .worksheet_parts(&reopened)
      .next()
      .expect("first worksheet");
    let xml = first_sheet
      .data_as_str(&reopened)
      .expect("worksheet xml")
      .expect("worksheet data");

    assert!(
      xml.contains(
        r#"<row r="4"><c r="C4" t="inlineStr"><is><t>New &amp; Value</t></is></c></row>"#
      )
    );
  }

  #[test]
  fn deletes_text_from_cell() {
    let fixture = write_spreadsheet_fixture();

    let bytes = delete_text_from_cell(&fixture, "Summary", "A2").expect("delete text");

    let reopened = SpreadsheetDocument::new(Cursor::new(bytes)).expect("reopen spreadsheet");
    let workbook_part = reopened.workbook_part().expect("workbook part");
    let first_sheet = workbook_part
      .worksheet_parts(&reopened)
      .next()
      .expect("first worksheet");
    let xml = first_sheet
      .data_as_str(&reopened)
      .expect("worksheet xml")
      .expect("worksheet data");

    assert!(xml.contains(r#"<c r="A2"/>"#));
    assert!(!xml.contains(r#"<c r="A2" t="s"><v>2</v></c>"#));
  }

  #[test]
  fn merges_adjacent_cells() {
    let fixture = write_spreadsheet_fixture();

    let bytes = merge_adjacent_cells(&fixture, "Summary", "A2", "B2").expect("merge cells");

    let reopened = SpreadsheetDocument::new(Cursor::new(bytes)).expect("reopen spreadsheet");
    let workbook_part = reopened.workbook_part().expect("workbook part");
    let first_sheet = workbook_part
      .worksheet_parts(&reopened)
      .next()
      .expect("first worksheet");
    let xml = first_sheet
      .data_as_str(&reopened)
      .expect("worksheet xml")
      .expect("worksheet data");

    assert!(xml.contains(r#"<mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>"#));
  }

  #[test]
  fn inserts_new_worksheet() {
    let fixture = write_spreadsheet_fixture();

    let bytes = insert_new_worksheet(&fixture, "Added Sheet").expect("insert worksheet");

    let reopened = SpreadsheetDocument::new(Cursor::new(bytes)).expect("reopen spreadsheet");
    let workbook_part = reopened.workbook_part().expect("workbook part");
    let workbook_xml = workbook_part
      .data_as_str(&reopened)
      .expect("workbook xml")
      .expect("workbook data");

    assert_eq!(workbook_part.worksheet_parts(&reopened).count(), 3);
    assert!(workbook_xml.contains(r#"name="Added Sheet""#));
    assert!(workbook_xml.contains(r#"sheetId="3""#));
  }

  fn write_spreadsheet_fixture() -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!(
      "ooxmlsdk-doc-spreadsheet-{}-{}.xlsx",
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
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    )
    .expect("write package rels");

    zip.add_directory("xl", options).expect("xl dir");
    zip.add_directory("xl/_rels", options).expect("xl rels dir");
    zip
      .add_directory("xl/worksheets", options)
      .expect("worksheets dir");

    zip
      .start_file("xl/workbook.xml", options)
      .expect("workbook");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden Data" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="SalesRange">Summary!$B$2:$B$2</definedName>
  </definedNames>
</workbook>"#,
    )
    .expect("write workbook");

    zip
      .start_file("xl/_rels/workbook.xml.rels", options)
      .expect("workbook rels");
    zip.write_all(
      br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#,
    )
    .expect("write workbook rels");

    zip
      .start_file("xl/sharedStrings.xml", options)
      .expect("shared strings");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Region</t></si>
  <si><t>Sales</t></si>
  <si><t>North</t></si>
</sst>"#,
      )
      .expect("write shared strings");

    zip
      .start_file("xl/worksheets/sheet1.xml", options)
      .expect("sheet1");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="3" max="4" hidden="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>42</v></c>
    </row>
    <row r="3" hidden="1"/>
  </sheetData>
</worksheet>"#,
      )
      .expect("write sheet1");

    zip
      .start_file("xl/worksheets/sheet2.xml", options)
      .expect("sheet2");
    zip
      .write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" state="hidden">
  <sheetData/>
</worksheet>"#,
      )
      .expect("write sheet2");

    zip.finish().expect("finish fixture");
    path
  }
}
