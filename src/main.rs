use std::io;
// use serde::{ Serialize };
use calamine::{ open_workbook, Xlsx, Reader, DataType };
use csv::Writer;

// cargo lib: https://crates.io/crates/calamine

#[derive(Debug)]
struct Row {
    first: String,
    last: String,
    email: String,
    phone: String,
    license_id: String,
    designation: String,
}

fn main() {
    // get file name from user
    println!("Name of the file (example.xlsx):");
    let mut file_name = match get_input() {
        Ok(s) => s,
        Err(()) => {
            println!("Could not read input");
            let _ = get_input();
            return;
        }
    };

    // return if no input
    if file_name.len() < 1 {
        println!("Nothin to do");
        let _ = get_input();
        return;
    }

    // append file extension if none (fails for: ../test.xlsx)
    let text_vec: Vec<&str> = file_name.split('.').collect();
    if text_vec.len() < 2 {
        file_name = file_name + ".xlsx"
    }

    // open workbook
    let mut workbook: Xlsx<_> = match open_workbook(&file_name) {
        Ok(xlsx) => xlsx,
        Err(_err) => {
            println!("ERR: File \"{}\" not found", &file_name);
            let _ = get_input();
            return;
        }
    };

    println!("\nFound file \"{}\" - Opening Sheet1\n", &file_name);

    // collectors
    let mut page_collection: Vec<Vec<Row>> = Vec::new();
    let mut cur_page: Vec<Row> = Vec::new();
    let mut read_next_line: bool = false;

    // go through sheet 1
    if let Some(Ok(r)) = workbook.worksheet_range("Sheet1") {
        for row in r.rows() {
            // ignore empty rows
            if row[0] == DataType::Empty {
                read_next_line = false;
                continue;
            }
            // start reading next line
            if row[0] == DataType::String("first".to_string()) {
                read_next_line = true;
                continue;
            }
            // println!("row {}: {:?}", i, row);
            // add row to collection
            if read_next_line && cur_page.len() < 50 {
                let new_row: Row = {Row {
                    first: row[0].as_string().unwrap(),
                    last: row[1].as_string().unwrap(),
                    email: row[2].as_string().unwrap(),
                    phone: row[3].as_string().unwrap(),
                    license_id: row[4].as_string().unwrap(),
                    designation: row[5].as_string().unwrap(),
                }};
                cur_page.push(new_row);
            }
            // change page
            if read_next_line && cur_page.len() == 50 {
                page_collection.push(cur_page);
                cur_page = Vec::new();
                let new_row: Row = {Row {
                    first: row[0].as_string().unwrap(),
                    last: row[1].as_string().unwrap(),
                    email: row[2].as_string().unwrap(),
                    phone: row[3].as_string().unwrap(),
                    license_id: row[4].as_string().unwrap(),
                    designation: row[5].as_string().unwrap(),
                }};
                cur_page.push(new_row);
            }
        }
        page_collection.push(cur_page);
    } else {
        println!("Could not find Sheet1");
        let _ = get_input();
        return;
    }

    // println!("Collected pages: {:?}\n", page_collection);
    for (i, page) in page_collection.iter().enumerate() {
        println!("Page {}: {} rows", i+1, page.len());
        let write_path = "output_".to_owned() + &(i+1).to_string() + ".csv";
        let mut writer = Writer::from_path(write_path).unwrap();

        // add headers
        let _ = writer.write_record(&["first","last","email","phone","license id","designation"]);

        // add rows
        for row in page {
            let _ = writer.write_record([
                &row.first,
                &row.last,
                &row.email,
                &row.phone,
                &row.license_id,
                &row.designation
            ]);
        }
        let _ = writer.flush();
    }

    println!("\nFinished! Press Enter to exit");
    let _ = get_input();
}

fn get_input() -> Result<String, ()> {
    let mut buffer = String::new();
    let result = io::stdin().read_line(&mut buffer);

    // handle error
    match result {
        Err(_) => return Err(()),
        _ => (),
    };

    // remove end line
    buffer = buffer.trim_end().to_string();

    Ok(buffer)
}