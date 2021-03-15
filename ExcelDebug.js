/*const newFilename = process.env.npm_config_newfilename;
console.log(newFilename);*/

const fs = require('fs');

const reader = require('xlsx');

const jsonFileName = "./debug.json";
const file = reader.readFile("./generic_multiple.xlsx");

let data = [];

const sheets = file.SheetNames;

let debugJSON = [];

function createHeaders(headers) {
    let headerObjects = [];

    for (let index = 0; index < headers.length; index++) {
        let tempHeader = headers[index];
        
        let headerObj = {
            headerOriginalText: tempHeader.trim(),
            headerJSONLabel: castToPascalCase(tempHeader),
            headerIndex: index,
            ignoreFlag: false
        }

        headerObjects.push(headerObj);
    }

    return headerObjects;
}

function castToPascalCase(val) {
    return val.trim().toLowerCase().replace(/ /g, "_");
}

for(let i = 0; i < sheets.length; i++) {
    const sheetname = file.SheetNames[i];

    const newObj = {};
    newObj["SheetName"] = sheetname;
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], {header: 1});

    // IGNORE FIRST ROW
    newObj["NodesCount"] = temp.length - 1;

    const worksheet = file.Sheets[file.SheetNames[i]];

    debugJSON.push(worksheet);

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
fs.writeFile(jsonFileName, jsonStr, function (err) {
    if (err) return console.log(err);
    console.log('Completed Excel to JSON > ' + jsonFileName);
});