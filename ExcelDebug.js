/*const newFilename = process.env.npm_config_newfilename;
console.log(newFilename);*/


const fs = require('fs');

const reader = require('xlsx');

const jsonFileName = "./debug.json";
const file = reader.readFile("./dummy.xlsx");

const debugFileName = "./debugging.txt";

let data = [];

const sheets = file.SheetNames;

let debugJSON = [];

for(let i = 0; i < sheets.length; i++) {
    const sheetname = file.SheetNames[i];

    const newObj = {};
    newObj["SheetName"] = sheetname;

    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], {header: 1});

    const worksheet = file.Sheets[file.SheetNames[i]];
    //console.log(temp);
    //console.log(temp[0]);
    //console.log(worksheet);
    //console.table(temp);
    //console.table(worksheet);

    debugJSON.push(worksheet);

    newObj.data = [];

    temp.forEach((res) => {
        const newRow = [];

        newRow[res[0]] = res[0];

        newObj.data.push(newRow);
        //console.log(res);
        console.log(newRow);
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

jsonStr = JSON.stringify(debugJSON);

fs.writeFile(debugFileName, jsonStr, function (err) {
    if (err) return console.log(err);
    console.log('Completed Excel to JSON > ' + jsonFileName);
});
