const fs = require("fs");
const md5 = require("md5");
const path = require("path");
const xlsx = require("xlsx");
const unzip = require("unzipper");
const {XMLParser} = require('fast-xml-parser');

const test_path = "source/test.xlsx";
const test_name = /^(.*)\/(\w+)\.xlsx$/.exec(test_path)[2];
const test_source_dir = path.join(__dirname, `source/${test_name}_resources`).toString();

// Типы тестов
const test_types = {
    OO: "one_option",
    OOMQ: "one_option_many_question",
    MO: "many_option",
    COS: "correct_option_sequence"
};

// Уровни сложности вопросов
const question_lvls = {
    base: "base",
    middle: "middle",
    hard: "hard"
};

// Функция парса таблицы в тест
const parse_table = (name, path_table, path_source) => {
    const workbook = xlsx.readFile(path_table);
    const sheet_name_list = workbook.SheetNames;
    const sheet_data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

    const xml_parser = new XMLParser({
        ignoreAttributes : false
    });

    let xml_drawing_structure = xml_parser.parse(fs.readFileSync(path.join(path_source, "xl", "drawings", "drawing1.xml")));
    let xml_drawing_rels_structure = xml_parser.parse(fs.readFileSync(path.join(path_source, "xl", "drawings", "_rels", "drawing1.xml.rels")));

    const normalize_table = [];

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
            
            const rel_image_name = /^(.*)\/(.*)$/.exec(rel["@_Target"])[2];
            rels_images[rel["@_Id"]] = {
                name: rel_image_name,
                path: path.join(path_source, "xl", "media", rel_image_name),
            };
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

    // Создание необходимых директорий 
    const path_save_tests = path.join(__dirname, "tests");
    if(!fs.existsSync(path_save_tests))
        fs.mkdirSync(path_save_tests)

    const test_save_root_path = path.join(path_save_tests, name);
    if(!fs.existsSync(test_save_root_path))
        fs.mkdirSync(test_save_root_path)

    const test_save_media_path = path.join(test_save_root_path, "media");
    if(!fs.existsSync(test_save_media_path))
        fs.mkdirSync(test_save_media_path)

    // Структура теста
    const test_structure = {
        theme: sheet_data[0]["__EMPTY_2"],
        list: []
    };

    console.log(normalize_table);

    // Парсим нормализированную таблицу в новый вид
    for(let rId = 0; rId < normalize_table.length; rId++) {
        const current_row = normalize_table[rId];

        // Записываем необходимые данные о вопросе
        const type = current_row["2"].text;
        const theme = current_row["3"].text;
        const lvl = current_row["4"].text;
        const is_milestone = (current_row["8"].text === 1) ? true : false;
        const text = current_row["10"].text;

        // Ответы
        const answers = [];
        const questions = [];

        // Данные о тесте
        let test_data;

        // Индекс начала просчётов
        const column_option_start_index = 11;

        // Далее в зависимости от типа
        switch(type) {
            // Один вариант вопроса, несколько вариантов ответа
            case test_types.OO: 
                for(let elId = column_option_start_index + 1; elId < Object.keys(current_row).length; elId += 2) {
                    const el = current_row[elId.toString()];
                    let isCorrect = false;
                    
                    // Если это первый ответ - значит он верный
                    if(column_option_start_index === elId)
                        isCorrect = true;

                    // Создаём объект ответа
                    const data = {
                        text: el.text,
                        isCorrect
                    };

                    // Экспортируем изображения (если имеются)
                    if(el.image) {
                        const image_path = path.join(test_save_media_path, el.image.name);
                        fs.renameSync(el.image.path, image_path);
                        data["image"] = image_path;
                    }

                    // Заносим ответ на вопрос
                    answers.push(data);
                }

                test_data = {
                    type,
                    theme,
                    lvl,
                    is_milestone,
                    text,
                    answers,
                };
                break;

            // Один вопрос, несколько возможных вариантов ответа
            case test_types.MO: 
                let el_number = 1;
                for(let elId = column_option_start_index + 1; elId < Object.keys(current_row).length; elId += 2) {
                    const el = current_row[elId.toString()];
                    const correct_number = current_row[column_option_start_index.toString()].text;
                    let isCorrect = false;

                    // Если это первый ответ - значит он верный
                    if(el_number === correct_number)
                        isCorrect = true;

                    // Создаём объект ответа
                    const data = {
                        text: el.text,
                        isCorrect
                    };

                    // Экспортируем изображения (если имеются)
                    if(el.image) {
                        const image_path = path.join(test_save_media_path, el.image.name);
                        fs.renameSync(el.image.path, image_path);
                        data["image"] = image_path;
                    }

                    // Заносим ответ на вопрос
                    answers.push(data);

                    el_number++;
                }

                test_data = {
                    type,
                    theme,
                    lvl,
                    is_milestone,
                    text,
                    answers,
                };
                break;

            // Несколько вариантов вопроса, несколько вариантов ответа
            case test_types.OOMQ:
            case test_types.COS: 
                // Добавление ответов
                for(let elId = column_option_start_index + 1; elId < Object.keys(current_row).length; elId += 2) {
                    const el = current_row[elId.toString()];

                    // Создаём объект ответа
                    const data = {
                        text: el.text
                    };

                    // Экспортируем изображения (если имеются)
                    if(el.image) {
                        const image_path = path.join(test_save_media_path, el.image.name);
                        fs.renameSync(el.image.path, image_path);
                        data["image"] = image_path;
                    }

                    // Заносим ответ на вопрос
                    answers.push(data);
                }
                
                // Добавление вопросов
                let question_index = 0;
                for(let elId = column_option_start_index; elId < Object.keys(current_row).length; elId += 2) {
                    const el = current_row[elId.toString()];

                    // Создаём объект вопроса
                    const data = {
                        text: el.text,
                        correctAnswerIndex: question_index
                    };

                    // Экспортируем изображения (если имеются)
                    if(el.image) {
                        const image_path = path.join(test_save_media_path, el.image.name);
                        fs.renameSync(el.image.path, image_path);
                        data["image"] = image_path;
                    }

                    // Заносим вопрос
                    questions.push(data);

                    question_index++;
                }

                test_data = {
                    type,
                    theme,
                    lvl,
                    is_milestone,
                    text,
                    answers,
                    questions
                };
                break;

            // Верная последовательность единиц
            // case test_types.COS: 
            //     // Добавление ответов
            //     for(let elId = column_option_start_index + 1; elId < Object.keys(current_row).length; elId += 2) {
            //         const el = current_row[elId.toString()];

            //         // Создаём объект ответа
            //         const data = {
            //             text: el.text
            //         };

            //         // Экспортируем изображения (если имеются)
            //         if(el.image) {
            //             const image_path = path.join(test_save_media_path, el.image.name);
            //             fs.renameSync(el.image.path, image_path);
            //             data["image"] = image_path;
            //         }

            //         // Заносим ответ на вопрос
            //         answers.push(data);
            //     }
                
            //     // Добавление вопросов
            //     let question_index = 0;
            //     for(let elId = column_option_start_index; elId < Object.keys(current_row).length; elId += 2) {
            //         const el = current_row[elId.toString()];

            //         // Создаём объект вопроса
            //         const data = {
            //             text: el.text,
            //             correctAnswer: answers[question_index]
            //         };

            //         // Экспортируем изображения (если имеются)
            //         if(el.image) {
            //             const image_path = path.join(test_save_media_path, el.image.name);
            //             fs.renameSync(el.image.path, image_path);
            //             data["image"] = image_path;
            //         }

            //         // Заносим вопрос
            //         questions.push(data);

            //         question_index++;
            //     }

            //     test_data = {
            //         type,
            //         theme,
            //         lvl,
            //         is_milestone,
            //         text,
            //         answers,
            //         questions
            //     };
            //     break;
        }

        // Добавить новый вопрос в лист теста
        test_structure.list.push(test_data);
    }

    // Записываем структуру в файл
    fs.writeFileSync(path.join(test_save_root_path, "structure.json"), JSON.stringify(test_structure));
    // Удаляем лишний мусор после работы
    fs.rmSync(path_source, { recursive: true });
}

// Разархивируем XLSX таблицу
fs
    .createReadStream(test_path)
    .pipe(unzip.Extract({path: test_source_dir}))
    .on('close', () => {
        parse_table(test_name, test_path, test_source_dir);
    });