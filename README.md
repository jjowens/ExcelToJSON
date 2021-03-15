# ExcelToJSON
Exporting Excel Worksheets to JSON. It will use your row headers as node names and cell values as your data. 

## Requirements

You are required to have the following npm packaage.

xlsx - https://www.npmjs.com/package/xlsx

## NPM command

```npm run exportjson``` is the main command for export a spreadsheet to a JSON file. You may be required to use the following parameters.

The data will be split into each node. Each node will use the sheetname as the name for that node. It requires a row header. When passing in a command parameter, case sensitivty does not matter and white spaces will be trimmed at start and end of the string values. All parameters are converted to lower case and trimmed for logic checks such as validating sheet names or row headers.


## Parameters

```--filepath=SET_FILEPATH_TO_SPREADSHEET``` - Required. Setting the "filepath" instructs the script to read the spreadsheet from the provided filepath.

```--jsonfilepath=SET_FILEPATH_TO_EXPORT_JSON``` - Required. Setting the "jsonfilepath" instructs the script to set the filepath to export the json data. If the "jsonfilepath" is not set, it will default to "data.json".

```--ignoresheets=ADD_SHEETNAMES_WITH_DELIMITERS``` - Optional. Add sheetnames to instruct the script to ignore sheets. Not case sensitive. Separate sheetsnames with a semi-colon e.g. ```--ignoresheets="Sheet 2; Sheet 6; Sheet 42"```

```--ignoreheaders=ADD_HEADERS_WITH_DELIMITERS``` - Optional. If you want the script to ignore certain columns, add header names. Not case sensitive. Separate header names with a semi-colon e.g. ```--ignoreheaders="Unit Price; Publisher Contact Name; Publisher Contact Number"```

## Examples

You can use spreadsheets, books.xlsx or books_stock.xlsx, to test the npm command to get various outputs.

### Basic

```npm run exportjson --filepath=./books.xlsx --jsonfilepath=data.json``` 

This will export data from one worksheet to data.json

```npm run exportjson --filepath=./books_stock.xlsx --jsonfilepath=data.json``` 

This will export data from all worksheets to data.json. Each node will be labelled with the sheetname.

### Advanced

```npm run exportjson --filepath=./books.xlsx --jsonfilepath=data.json --ignoreheaders="Genres"```

This will export data from one worksheet to data.json. It will ignore the column "Genres". It won't be part of the export.

```npm run exportjson --filepath=./books_stock.xlsx --jsonfilepath=data_stock.json --ignoreheaders="Genres; Publisher"```

This will export data from all worksheets to a file data_stock.json. It will ignore the columns with the row headers "Genres" and "Publisher". Those ignored columns won't be part of the export.

```npm run exportjson --filepath=./books_stock.xlsx --jsonfilepath=data_stock.json --ignoreheaders="Genres; Publisher"```

This will export data from all worksheets to a file data_stock.json. It will ignore the columns with the row headers "Genres" and "Publisher". Those ignored columns won't be part of the export.

```npm run exportjson --filepath=./books_stock.xlsx --jsonfilepath=data_stock.json --ignoreheaders="Genres; Publisher" --ignoresheets="Edinburgh Branch"```

This will export data from most worksheets to a file data_stock.json. It will ignore the "Edinburgh Branch" worksheet and then it will ignore columns with the row header "Genres" and "Publisher" from any other worksheets. Those ignoreed worksheets and columns won't be part of the export.