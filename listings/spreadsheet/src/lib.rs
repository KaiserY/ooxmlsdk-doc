// ANCHOR: open_spreadsheet_read_only
use std::path::Path;

use ooxmlsdk::parts::spreadsheet_document::SpreadsheetDocument;
use ooxmlsdk::sdk::{OpenSettings, PackageOpenMode};

pub fn open_spreadsheet_read_only(path: &Path) -> Result<usize, Box<dyn std::error::Error>> {
  let document = SpreadsheetDocument::new_from_file_with_settings(path, lazy_settings())?;
  let workbook_part = document.workbook_part()?;

  Ok(workbook_part.worksheet_parts(&document).count())
}
// ANCHOR_END: open_spreadsheet_read_only

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
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2"><v>42</v></c>
    </row>
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
