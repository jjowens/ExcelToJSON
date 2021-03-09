const fs = require('fs');

const reader = require('xlsx');

const jsonFileName = "./data.json";
const file = reader.readFile("./generic.xlsx");

let data = [];

const sheets = file.SheetNames;

for(let i = 0; i < sheets.length; i++) {
    const sheetname = file.SheetNames[i];

    const newObj = {};
    newObj["SheetName"] = sheetname;

    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);

    const worksheet = file.Sheets[file.SheetNames[i]];
    console.table(worksheet.keys);

    newObj.data = [];

    temp.forEach((res) => {
        const newRow = [];

        newRow[res[0]] = res[0];

        newObj.data.push(newRow);
    });

    data.push(newObj);
}

let jsonStr = JSON.stringify(data);

//console.log(jsonStr);

// WRITE JSON DATA TO FILE
fs.writeFile(jsonFileName, jsonStr, function (err) {
    if (err) return console.log(err);
    console.log('Completed Excel to JSON > ' + jsonFileName);
});
