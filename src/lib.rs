use std::path::Path;

use io::toml_reader::AppOptions;
use polars::prelude::*;

pub mod io;

const ALL_COMPONENT_NAME: &str = "all component functions";
const COLUMN_CENTER: &str = "Center";

fn fetch_peaks_df(filename: impl AsRef<Path>) -> PolarsResult<DataFrame> {
    let peaks_df = io::read_as_data_frame(filename, true, '\t')?;

    Ok(peaks_df)
}

fn fetch_fit_df(
    filename: impl AsRef<Path>,
    peaks_df: &DataFrame,
    option_col: &[impl AsRef<str>],
) -> PolarsResult<DataFrame> {
    let series_center = peaks_df.column(COLUMN_CENTER)?;
    let center_names = series_center
        .iter()
        .map(|name| name.to_string())
        .collect::<Vec<String>>();
    let sorted_center_names = series_center
        .sort(SortOptions::default())?
        .iter()
        .map(|name| name.to_string())
        .collect::<Vec<String>>();

    let dat_df = {
        let mut df = io::read_as_data_frame(filename, false, ' ')?;
        option_col
            .iter()
            .flat_map(|name| {
                if name.as_ref() == ALL_COMPONENT_NAME {
                    center_names.clone()
                } else {
                    vec![name.as_ref().to_string()]
                }
            })
            .enumerate()
            .for_each(|(index, name)| {
                df.rename(&format!("column_{}", index + 1), &name).unwrap();
            });

        let sort_option = option_col.iter().flat_map(|name| {
            if name.as_ref() == ALL_COMPONENT_NAME {
                sorted_center_names.clone()
            } else {
                vec![name.as_ref().to_string()]
            }
        });

        df.select(sort_option)?
    };

    Ok(dat_df)
}

pub fn sort_peaks(
    filename: &str,
    fityk_options: &[impl AsRef<str>],
    app_options: &AppOptions,
) -> anyhow::Result<()> {
    let cd = std::env::current_dir()?;
    let read_folder = app_options.folder.read.as_ref().unwrap_or(&cd);
    let write_folder = app_options.folder.write.as_ref().unwrap_or(&cd);

    let peaks_file = read_folder.join(format!("{filename}.peaks"));
    let peaks_data = fetch_peaks_df(peaks_file)?;

    let dat_file = read_folder.join(format!("{filename}.dat"));
    let mut fit_data = fetch_fit_df(dat_file, &peaks_data, fityk_options)?;

    let xlsx_file = write_folder.join(format!("{filename}.xlsx"));
    io::write_to_excel(xlsx_file, &peaks_data, &fit_data)?;

    if app_options.create_csv {
        let mut csv_file = std::fs::File::create(write_folder.join(format!("CSV_{filename}.csv")))?;
        CsvWriter::new(&mut csv_file).finish(&mut fit_data)?
    }

    Ok(())
}
