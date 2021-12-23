const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const unzip = require("unzipper");
const {XMLParser} = require('fast-xml-parser');

const test_path = path.join(__dirname, "source/test.xlsx");
const test_name = /^(.*)\/(\w+)\.xlsx$/.exec(test_path)[2];
const test_source_dir = path.join(__dirname, `source/${test_name}_resources`).toString();

const parse_table = (path_table, path_source) => {
    const workbook = xlsx.readFile(path_table);
    const sheet_name_list = workbook.SheetNames;
    const sheet_data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

    const xml_parser = new XMLParser({
        ignoreAttributes : false
    });

    let xml_drawing_structure = xml_parser.parse(fs.readFileSync(path.join(path_source, "xl", "drawings", "drawing1.xml")));
    let xml_drawing_rels_structure = xml_parser.parse(fs.readFileSync(path.join(path_source, "xl", "drawings", "_rels", "drawing1.xml.rels")));

    const normalize_table = [];

    const tests = {
        theme: sheet_data[0]["__EMPTY_2"],
        list: []
    };

    // Нормализируем таблицу
    for(let i = 2; i < sheet_data.length; i++) {
        const sheet_row_data = sheet_data[i];
        const final_row = {};

        // Функция добавления столбца в строку
        const adding_row = (index, options) => {
            const row = final_row[index.toString()] || {};
            for(let key in options) {
                row[key] = options[key];
            }
            final_row[index.toString()] = row;
        }

        // Получаем ассоциации id с картинками
        let rels_images = {};
        for(let rel of xml_drawing_rels_structure["Relationships"]["Relationship"]) {
            rels_images[rel["@_Id"]] = /^(.*)\/(.*)$/.exec(rel["@_Target"])[2];
        }

        // Получаем все картинки для текущей строки
        for(let obj of xml_drawing_structure["xdr:wsDr"]["xdr:twoCellAnchor"]) {
            if(obj["xdr:pic"]) {
                let from = obj["xdr:from"];
                if((from["xdr:row"] - 4) === i) {
                    let image_id = obj["xdr:pic"]["xdr:blipFill"]["a:blip"]["@_r:embed"];
                    adding_row(from["xdr:col"], {image: rels_images[image_id]});
                }
            }
        }
        
        // Перебираем все столбцы строки
        for(let key in sheet_row_data) {
            let regex = /(.*)_(\d+)/;

            let index = 0;
            if (regex.test(key))
                index = regex.exec(key)[2];

            adding_row(index, {text: sheet_row_data[key]})
        }

        normalize_table.push(final_row);
    }

    console.log(normalize_table);
}

// Разархивируем XLSX таблицу
fs
    .createReadStream(test_path)
    .pipe(unzip.Extract({path: test_source_dir}))

parse_table(test_path, test_source_dir);