use core::panic;
use pyo3::prelude::*;
use rust_xlsxwriter::worksheet::{ColNum, RowNum};
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook, Worksheet};

fn border_style_from_index(index: u32) -> FormatBorder {
    match index {
        0 => FormatBorder::None,
        1 => FormatBorder::Thin,
        2 => FormatBorder::Medium,
        7 => FormatBorder::Hair,
        _ => panic!(),
    }
}

fn format_align_from_string(format_align: &str) -> FormatAlign {
    match format_align {
        "left" => FormatAlign::Left,
        "center" => FormatAlign::Center,
        "vcenter" => FormatAlign::VerticalCenter,
        _ => panic!(),
    }
}

#[derive(Debug)]
#[pyclass(subclass, unsendable)]
struct RustFormat {
    _format: Format,
}

#[pymethods]
impl RustFormat {
    #[new]
    fn new() -> Self {
        RustFormat {
            _format: Format::new(),
        }
    }
}

#[pyclass(subclass, unsendable)]
struct RustWorkbook {
    _workbook: Workbook,
    _worksheets: Vec<RustWorksheet>,
}

#[pymethods]
impl RustWorkbook {
    #[new]
    fn new() -> Self {
        RustWorkbook {
            _workbook: Workbook::new(),
            _worksheets: Vec::new(),
        }
    }

    fn add_worksheet(&mut self, name: String) -> PyResult<RustWorksheet> {
        let mut worksheet = RustWorksheet::new();
        worksheet.set_name(name)?;

        Ok(worksheet)
    }

    fn close(&mut self, filename: String, worksheets: Vec<PyRefMut<RustWorksheet>>) {
        for mut sheet in worksheets {
            // This works for the one worksheet but I still need to figure out a list of worksheets
            // we need to move the worksheet out of the RustWorksheet struct, see: https://stackoverflow.com/a/31308299
            let temp = std::mem::replace(&mut sheet._worksheet, Worksheet::new());
            self._workbook.push_worksheet(temp);
        }

        self._workbook.save(filename).unwrap();
    }

    // add format, dict can contain ints, strings and booleans
    #[pyo3(signature = (bold=None, font_name=None, font_size=None, font_color=None, align=None, valign=None, left=None, right=None, top=None, bottom=None, indent=None, bg_color=None, text_wrap=None, num_format=None))]
    fn add_format(
        &mut self,
        bold: Option<bool>,
        font_name: Option<String>,
        font_size: Option<u32>,
        font_color: Option<String>,
        align: Option<String>,
        valign: Option<String>,
        left: Option<u32>,
        right: Option<u32>,
        top: Option<u32>,
        bottom: Option<u32>,
        indent: Option<u32>,
        bg_color: Option<String>,
        text_wrap: Option<bool>,
        num_format: Option<String>,
    ) -> PyResult<RustFormat> {
        let mut format = RustFormat::new();
        if bold.is_some() && bold.unwrap() {
            format._format = format._format.set_bold();
        }

        if font_name.is_some() {
            format._format = format._format.set_font_name(font_name.unwrap());
        }

        if font_size.is_some() {
            format._format = format._format.set_font_size(font_size.unwrap());
        }

        if font_color.is_some() {
            format._format = format._format.set_font_color(font_color.unwrap().as_str());
        }

        if align.is_some() {
            format._format = format
                ._format
                .set_align(format_align_from_string(align.unwrap().as_str()));
        }

        if valign.is_some() {
            format._format = format
                ._format
                .set_align(format_align_from_string(valign.unwrap().as_str()));
        }

        if left.is_some() {
            format._format = format
                ._format
                .set_border_left(border_style_from_index(left.unwrap()));
        }

        if right.is_some() {
            format._format = format
                ._format
                .set_border_right(border_style_from_index(right.unwrap()));
        }

        if top.is_some() {
            format._format = format
                ._format
                .set_border_top(border_style_from_index(top.unwrap()));
        }

        if bottom.is_some() {
            format._format = format
                ._format
                .set_border_bottom(border_style_from_index(bottom.unwrap()));
        }

        if indent.is_some() {
            format._format = format._format.set_indent(indent.unwrap() as u8);
        }

        if bg_color.is_some() {
            format._format = format
                ._format
                .set_background_color(bg_color.unwrap().as_str());
        }

        // text_wrap: Option<bool>,
        // num_format: Option<String>,

        Ok(format)
    }
}

#[pyclass(subclass, unsendable)]
struct RustWorksheet {
    pub _worksheet: Worksheet,
}

#[pymethods]
impl RustWorksheet {
    #[new]
    fn new() -> Self {
        RustWorksheet {
            _worksheet: Worksheet::new(),
        }
    }

    #[setter]
    fn set_name(&mut self, name: String) -> PyResult<()> {
        self._worksheet.set_name(name).unwrap();

        Ok(())
    }

    #[getter]
    fn get_name(&self) -> PyResult<String> {
        Ok(self._worksheet.name())
    }

    fn set_column_width(&mut self, col: ColNum, width: u32) {
        self._worksheet.set_column_width(col, width).unwrap();
    }

    fn set_row_height(&mut self, row: RowNum, height: f64) {
        self._worksheet.set_row_height(row, height).unwrap();
    }

    fn set_default_row_height(&mut self, height: u32) {
        self._worksheet.set_default_row_height(height);
    }

    fn set_screen_gridlines(&mut self, enable: bool) {
        self._worksheet.set_screen_gridlines(enable);
    }

    fn set_print_gridlines(&mut self, enable: bool) {
        self._worksheet.set_print_gridlines(enable);
    }

    fn set_column_hidden(&mut self, col: ColNum) {
        self._worksheet.set_column_hidden(col).unwrap();
    }

    fn group_columns_collapsed(&mut self, col: ColNum) {
        self._worksheet.group_columns_collapsed(col, col).unwrap();
    }

    #[pyo3(signature = (row, col, format_obj=None))]
    fn write_blank(&mut self, row: RowNum, col: ColNum, format_obj: Option<PyRefMut<RustFormat>>) {
        if format_obj.is_some() {
            // println!("{:?}", format_obj);
            self._worksheet
                .write_blank(row, col, &format_obj.unwrap()._format)
                .unwrap();
        } else {
            let default_format = Format::new();
            self._worksheet
                .write_blank(row, col, &default_format)
                .unwrap();
        }
    }

    fn write_string(&mut self, row: RowNum, col: ColNum, data: String) {
        self._worksheet.write_string(row, col, data).unwrap();
    }

    fn write_string_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        data: String,
        format_obj: PyRefMut<RustFormat>,
    ) {
        self._worksheet
            .write_string_with_format(row, col, data, &format_obj._format)
            .unwrap();
    }

    fn write_formula(&mut self, row: RowNum, col: ColNum, formula: String) {
        self._worksheet
            .write_formula(row, col, formula.as_str())
            .unwrap();
    }

    fn write_formula_with_format(
        &mut self,
        row: RowNum,
        col: ColNum,
        formula: String,
        format_obj: PyRefMut<RustFormat>,
    ) {
        self._worksheet
            .write_formula_with_format(row, col, formula.as_str(), &format_obj._format)
            .unwrap();
    }

    fn merge_range(
        &mut self,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
        data: String,
        format_obj: PyRefMut<RustFormat>,
    ) {
        self._worksheet
            .merge_range(
                first_row,
                first_col,
                last_row,
                last_col,
                data.as_str(),
                &format_obj._format,
            )
            .unwrap();
    }
}

/// A Python module implemented in Rust.
#[pymodule]
fn xlsx_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<RustWorksheet>()?;
    m.add_class::<RustWorkbook>()?;
    Ok(())
}
