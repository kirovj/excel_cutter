use calamine::{open_workbook, DataType, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, XlsxError};
use std::error::Error;

fn push_header(datas: &mut Vec<Vec<String>>, header: &Vec<String>) {
    let mut _header = Vec::new();
    for h in header {
        _header.push(String::from(h));
    }
    datas.push(_header);
}

fn row_to_vec(row: &[DataType]) -> Vec<String> {
    let mut data = Vec::new();
    for ele in row {
        data.push(String::from(ele.get_string().unwrap_or("")));
    }
    data
}

fn process_excel(filepath: &str, name: &str, limit: usize) -> Result<(), Box<dyn Error>> {
    let mut excel: Xlsx<_> = open_workbook(filepath).expect("Cannot open file");

    // Read whole worksheet data and provide some statistics
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
        let mut datas: Vec<Vec<String>> = Vec::new();
        let mut header: Vec<String> = Vec::new();
        let mut num = 0;
        for (index, row) in r.rows().enumerate() {
            if index == 0 {
                header = row_to_vec(row);
                push_header(&mut datas, &header);
            } else if index % limit == 0 {
                let _ = write_excel(format!("{}_{}.xlsx", name, num), datas);
                num += 1;
                datas = Vec::new();
                push_header(&mut datas, &header);
            } else {
                datas.push(row_to_vec(row));
            }
        }
        if datas.len() > 0 {
            let _ = write_excel(format!("{}_{}.xlsx", name, num), datas);
        }
    }
    Ok(())
}

fn write_excel(filepath: String, datas: Vec<Vec<String>>) -> Result<(), XlsxError> {
    println!("start writing {}, {} rows", filepath, datas.len());
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let mut row = 0;
    for data in datas {
        let mut col = 0;
        for d in data {
            worksheet.write(row, col, d.as_str())?;
            col += 1;
        }
        row += 1;
    }
    workbook.save(filepath)?;
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let _limit = 3000;
    let path = "examples/a.xlsx";
    let path_string = String::from(path);
    let mut split = path_string.split(".");
    let _name = split.next().expect("filename is empty.");
    match split.next() {
        Some("xls") | Some("xlsx") => process_excel(path, _name, _limit),
        Some(_type) => {
            println!("file type `{}` is not support!", _type);
            Ok(())
        }
        _ => {
            println!("file type is empty.");
            Ok(())
        }
    }

    // opens a new workbook
}
