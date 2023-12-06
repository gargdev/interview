const fs = require('fs');
const XLSX = require('xlsx');

// Function to convert nested JSON to flat JSON
function flattenJSON(obj, parentKey = '') {
  let result = {};

  for (const key in obj) {
    const newKey = parentKey ? `${parentKey}_${key}` : key;

    if (typeof obj[key] === 'object' && obj[key] !== null) {
      result = { ...result, ...flattenJSON(obj[key], newKey) };
    } else {
      result[newKey] = obj[key];
    }
  }

  return result;
}

// Read JSON file
const jsonFilePath = './data/example.json'; // Replace with your JSON file path
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf-8'));

// Ensure jsonData is an array
const dataArray = Array.isArray(jsonData) ? jsonData : [jsonData];

// Create workbook and main sheet
const workbook = XLSX.utils.book_new();
const mainSheetData = dataArray.map((data) => flattenJSON(data));
const mainSheet = XLSX.utils.json_to_sheet(mainSheetData);
XLSX.utils.book_append_sheet(workbook, mainSheet, 'MainSheet');

// Create separate sheet for nested data
const nestedSheetData = dataArray.map((data, index) => {
  return {
    Index: index + 1,
    NestedData: JSON.stringify(data), // Adjust this line based on your JSON structure
  };
});
const nestedSheet = XLSX.utils.json_to_sheet(nestedSheetData);
XLSX.utils.book_append_sheet(workbook, nestedSheet, 'NestedDataSheet');

// Link the nested sheet to the main sheet
mainSheet['A1'].l = { r: 0, c: 1, t: 'sheet', $: 'NestedDataSheet' };

// Write the workbook to an Excel file
const excelFilePath = './output/output.xlsx'; // Replace with your desired output file path
XLSX.writeFile(workbook, excelFilePath);

console.log(`Excel file created at: ${excelFilePath}`);
