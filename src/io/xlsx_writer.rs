use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, TableStyle, Worksheet, XlsxError};

#[derive(Clone)]
pub struct WriterOptions {
    pub(crate) use_autofit: bool,
    pub(crate) float_format: Format,
    pub(crate) null_string: Option<String>,
    pub(crate) table: Option<Table>,
    pub(crate) zoom: u16,
    pub(crate) screen_gridlines: bool,
    pub(crate) freeze_cell: (u32, u16),
    pub(crate) top_cell: (u32, u16),
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl WriterOptions {
    fn new() -> WriterOptions {
        WriterOptions {
            use_autofit: true,
            null_string: None,
            float_format: Format::default(),
            table: Some(Table::new()),
            zoom: 100,
            screen_gridlines: true,
            freeze_cell: (0, 0),
            top_cell: (0, 0),
        }
    }
    pub fn with_use_autofit(mut self, use_autofit: bool) -> Self {
        self.use_autofit = use_autofit;
        self
    }
    pub fn with_null_string(mut self, null_string: Option<String>) -> Self {
        self.null_string = null_string;
        self
    }
    pub fn with_float_format(mut self, float_format: Format) -> Self {
        self.float_format = float_format;
        self
    }
    pub fn with_table(mut self, table: Option<Table>) -> Self {
        self.table = table;
        self
    }
    pub fn with_zoom(mut self, zoom: u16) -> Self {
        self.zoom = zoom;
        self
    }
    pub fn with_screen_gridlines(mut self, screen_gridlines: bool) -> Self {
        self.screen_gridlines = screen_gridlines;
        self
    }
    pub fn with_freeze_cell(mut self, freeze_cell: (u32, u16)) -> Self {
        self.freeze_cell = freeze_cell;
        self
    }
    pub fn with_top_cell(mut self, top_cell: (u32, u16)) -> Self {
        self.top_cell = top_cell;
        self
    }
    pub fn with_set_table_style(mut self, style: TableStyle) -> Self {
        if let Some(table) = self.table {
            self.table = Some(table.set_style(style));
        };

        self
    }
}

fn write_any_value(
    worksheet: &mut Worksheet,
    row_num: u32,
    col_num: u16,
    data: AnyValue,
    null_string: &Option<String>,
    float_format: &Format,
) -> Result<(), XlsxError> {
    match data {
        AnyValue::Int8(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::UInt8(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::Int16(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::UInt16(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::Int32(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::UInt32(value) => {
            worksheet.write_number(row_num, col_num, value)?;
        }
        AnyValue::Int64(value) => {
            // Allow u64 conversion within Excel's limits.
            #[allow(clippy::cast_precision_loss)]
            worksheet.write_number(row_num, col_num, value as f64)?;
        }
        AnyValue::UInt64(value) => {
            // Allow u64 conversion within Excel's limits.
            #[allow(clippy::cast_precision_loss)]
            worksheet.write_number(row_num, col_num, value as f64)?;
        }
        AnyValue::Float32(value) => {
            worksheet.write_number_with_format(row_num, col_num, value, float_format)?;
        }
        AnyValue::Float64(value) => {
            worksheet.write_number_with_format(row_num, col_num, value, float_format)?;
        }
        AnyValue::String(value) => {
            worksheet.write_string(row_num, col_num, value)?;
        }
        AnyValue::Boolean(value) => {
            worksheet.write_boolean(row_num, col_num, value)?;
        }
        AnyValue::Null => {
            if let Some(null_string) = null_string {
                worksheet.write_string(row_num, col_num, null_string)?;
            };
        }
        _ => {
            let message = format!(
                "Polars AnyValue data type '{}' is not supported by Excel",
                data.dtype()
            );
            return Err(XlsxError::CustomError(message));
        }
    };
    Ok(())
}

pub fn write_dataframe_internal(
    df: &DataFrame,
    worksheet: &mut Worksheet,
    row_offset: u32,
    col_offset: u16,
    options: &WriterOptions,
) -> Result<(), XlsxError> {
    let has_header = options.table.is_none() || options.table.as_ref().unwrap().has_header_row();
    let header_offset = u32::from(has_header);

    // Iterate through the dataframe column by column.
    for (col_num, column) in df.get_columns().iter().enumerate() {
        let col_num = col_offset + col_num as u16;

        // Store the column names for use as table headers.
        if has_header {
            worksheet.write(row_offset, col_num, column.name())?;
        }

        // Write the row data for each column/type.
        for (row_num, data) in column.iter().enumerate() {
            let row_num = header_offset + row_offset + row_num as u32;

            // Map the Polars Series AnyValue types to Excel/rust_xlsxwriter
            // types.
            write_any_value(
                worksheet,
                row_num,
                col_num,
                data,
                &options.null_string,
                &options.float_format,
            )?;
        }
    }

    // Add the table to the worksheet.
    if let Some(table) = &options.table {
        // Create a table for the dataframe range.
        let (mut max_row, max_col) = df.shape();
        if !table.has_header_row() {
            max_row -= 1;
        }
        if table.has_total_row() {
            max_row += 1;
        }

        worksheet.add_table(
            row_offset,
            col_offset,
            row_offset + max_row as u32,
            col_offset + max_col as u16 - 1,
            table,
        )?;
    }

    // Autofit the columns.
    if options.use_autofit {
        worksheet.autofit();
    }

    // Set the zoom level.
    worksheet.set_zoom(options.zoom);

    // Set the screen gridlines.
    worksheet.set_screen_gridlines(options.screen_gridlines);

    // Set the worksheet panes.
    worksheet.set_freeze_panes(options.freeze_cell.0, options.freeze_cell.1)?;
    worksheet.set_freeze_panes_top_cell(options.top_cell.0, options.top_cell.1)?;

    Ok(())
}
