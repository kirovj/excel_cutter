use calamine::{open_workbook, DataType, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, XlsxError};
use std::env;
use std::error::Error;
use std::thread;

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
    println!("读取 {}...", filepath);
    let mut excel: Xlsx<_> = open_workbook(filepath).expect("打开文件失败");
    let mut handlers = Vec::new();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
        println!("开始切割 {}...", filepath);
        let mut datas: Vec<Vec<String>> = Vec::new();
        let mut header: Vec<String> = Vec::new();
        let mut num = 0;

        for (index, row) in r.rows().enumerate() {
            if index == 0 {
                header = row_to_vec(row);
                push_header(&mut datas, &header);
            } else {
                datas.push(row_to_vec(row));
                if index % limit == 0 {
                    let _name = String::from(name);
                    handlers.push(thread::spawn(move || {
                        write_excel(format!("{}_{}.xlsx", _name, num), datas)
                    }));
                    num += 1;
                    datas = Vec::new();
                    push_header(&mut datas, &header);
                }
            }
        }
        if datas.len() > 0 {
            let _name = String::from(name);
            handlers.push(thread::spawn(move || {
                write_excel(format!("{}_{}.xlsx", _name, num), datas)
            }));
        }
    }
    for handler in handlers {
        let _ = handler.join().unwrap();
    }
    println!("切割完成！");
    Ok(())
}

fn write_excel(filepath: String, datas: Vec<Vec<String>>) -> Result<(), XlsxError> {
    println!("生成 {}, {} 行", filepath, datas.len() - 1);
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
    let args: Vec<String> = env::args().collect();
    assert!(
        args.len() == 3,
        "参数不够， example: `excel_cutter.exe a.xlsx 3000`"
    );

    let path = args[1].as_str();
    let _limit: usize = args[2].parse().expect("行数应该是数字");
    let path_string = String::from(path);
    let mut split = path_string.split(".");
    let _name = split.next().expect("文件名为空");
    match split.next() {
        Some("xls") | Some("xlsx") => process_excel(path, _name, _limit),
        Some(_type) => {
            println!("文件类型 `{}` 不支持", _type);
            Ok(())
        }
        _ => {
            println!("文件类型为空");
            Ok(())
        }
    }
}
