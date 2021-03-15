const filePath = process.env.npm_config_filepath;
let jsonFilePath = "data.json";
let ignoreHeaders = [];
let ignoreSheets = [];

if (filePath === undefined || "") {
    console.log("ERROR! 'filepath' parameter to spreadsheet is not provided");
    return;
}

if (process.env.npm_config_jsonfilepath !== undefined) {
    jsonFilePath = process.env.npm_config_jsonfilepath;
}

const fs = require('fs');
const reader = require('xlsx');

const file = reader.readFile(filePath);

let data = [];

const sheets = file.SheetNames;

function createHeaders(headers) {
    let headerObjects = [];

    for (let index = 0; index < headers.length; index++) {
        let tempHeader = headers[index];
        
        let headerObj = {
            headerOriginalText: tempHeader.trim(),
            headerLowerCase: castToLowerCase(tempHeader),
            headerJSONLabel: castToPascalCase(tempHeader),
            headerIndex: index
        }

        headerObjects.push(headerObj);
    }

    return headerObjects;
}

function castToPascalCase(val) {
    return val.trim().toLowerCase().replace(/ /g, "_");
}

function castToLowerCase(val) {
    return val.trim().toLowerCase();
}

// SET IGNORE HEADERS
if (process.env.npm_config_ignoreheaders !== undefined) {
    let tempIgnoreHeaders = process.env.npm_config_ignoreheaders.split(";");

    for (let index = 0; index < tempIgnoreHeaders.length; index++) {
        let tempIgnore = tempIgnoreHeaders[index];
        
        let tempIgnoreHeader = {
            headerText: tempIgnore.trim(),
            headerLowerCase: castToLowerCase(tempIgnore)
        }

        ignoreHeaders.push(tempIgnoreHeader);
    }

}

// SET IGNORE SHEETS
if (process.env.npm_config_ignoresheets !== undefined) {
    let tempIgnoreSheets = process.env.npm_config_ignoresheets.split(";");

    for (let index = 0; index < tempIgnoreSheets.length; index++) {
        let tempIgnore = tempIgnoreSheets[index];
        
        let tempIgnoreSheet = {
            sheetName: tempIgnore.trim(),
            sheetNameLowerCase: castToLowerCase(tempIgnore),
        }

        ignoreSheets.push(tempIgnoreSheet);
    }

}

for(let i = 0; i < sheets.length; i++) {
    const sheetname = file.SheetNames[i];
    const sheetnameLowercase = castToLowerCase(sheetname);

    // CHECK IF THIS WORKSHEET SHOULD BE IGNORED
    if (ignoreSheets.length > 0) {
        let ignoreFlag = ignoreSheets.some(item => item.sheetNameLowerCase == sheetnameLowercase);

        if (ignoreFlag === true) {
            continue;
        }
    }

    const newObj = {};
    newObj["name"] = sheetname;
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], {header: 1});

    // IGNORE FIRST ROW
    newObj["count"] = temp.length - 1;

    newObj.data = [];

    let headerObjs = createHeaders(temp[0]);

    for (let rowIndex = 0; rowIndex < temp.length; rowIndex++) {
        // IGNORE FIRST ROW HEADER
        if (rowIndex === 0) {
            continue;
        }
        
        let row = temp[rowIndex];

        let record = {};

        for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
            // CHECK IF THIS COLUMN SHOULD BE IGNORED
            if (ignoreHeaders.length > 0) {
                let ignoreFlag = ignoreHeaders.some(item => item.headerLowerCase == headerObjs[columnIndex].headerLowerCase);

                if (ignoreFlag === true) {
                    continue;
                }
            }
            
            const column = row[columnIndex];
            const columnTitle = headerObjs[columnIndex].headerJSONLabel;

            record[columnTitle] = column;
        }

        newObj.data.push(record);

    }

    data.push(newObj);
}


let jsonStr = JSON.stringify(data);

// WRITE JSON DATA TO FILE
fs.writeFile(jsonFilePath, jsonStr, function (err) {
    if (err) return console.log(err);
    console.log('Completed converting Excel to JSON > ' + jsonFilePath);
});