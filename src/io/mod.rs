pub mod toml_reader;
pub mod xlsx_writer;

use std::path::Path;

use polars::prelude::*;
use rust_xlsxwriter::{TableStyle, Workbook, XlsxError};
use toml_reader::{toml_to_settings, Settings};

pub fn read_as_data_frame(
    path: impl AsRef<Path>,
    has_header: bool,
    separator: char,
) -> PolarsResult<DataFrame> {
    let parse_option = CsvParseOptions::default().with_separator(separator as u8);

    let df = CsvReadOptions::default()
        .with_has_header(has_header)
        .with_parse_options(parse_option)
        .try_into_reader_with_file_path(Some(path.as_ref().to_path_buf()))?
        .finish()?
        .drop_nulls::<String>(None);

    df
}

pub fn write_to_excel(
    path: impl AsRef<Path>,
    peaks_data: &DataFrame,
    fit_data: &DataFrame,
) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let option = xlsx_writer::WriterOptions::default()
        .with_use_autofit(true)
        .with_set_table_style(TableStyle::None);

    xlsx_writer::write_dataframe_internal(
        &peaks_data,
        workbook.add_worksheet().set_name("PeaksData")?,
        0,
        0,
        &option,
    )?;

    let option = option.with_table(None);
    xlsx_writer::write_dataframe_internal(
        &fit_data,
        workbook.add_worksheet().set_name("FitData")?,
        0,
        0,
        &option,
    )?;

    workbook.save(path)?;

    Ok(())
}

pub fn read_toml_file(path: impl AsRef<Path>) -> anyhow::Result<Settings> {
    let file = std::fs::read_to_string(path)?;
    let settings = toml_to_settings(&file)?;

    Ok(settings)
}
