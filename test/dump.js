const reader = require('xlsx');
const fs = require('fs');
const path = require('path');

// Reading the Excel file
const file = reader.readFile('file.xlsx');

let data = [];

// Getting all sheet names
const sheets = file.SheetNames;

// Loop through all sheets and convert each to JSON
for (let i = 0; i < sheets.length; i++) {
  const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
  temp.forEach((res) => {
    data.push(res);
  });
}

// Define the path to save the JSON file
const outputPath = path.join(__dirname, 'dumpData.json');

// Save the data in a JSON file
fs.writeFileSync(outputPath, JSON.stringify(data, null, 2), 'utf-8');

console.log(`Excel data has been saved to: ${outputPath}`);
