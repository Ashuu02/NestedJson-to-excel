const fs = require('fs');
const XLSX = require('xlsx');

function convertNestedJsonToExcel(jsonData, excelFileName) {
  const wb = XLSX.utils.book_new();

  // Flatten the entire JSON object into the first sheet
  const flattenedData = flattenObject(jsonData);
  const wsAll = XLSX.utils.json_to_sheet([flattenedData], { header: Object.keys(flattenedData) });
  XLSX.utils.book_append_sheet(wb, wsAll, 'AllData');

  // Create sheets for each nested object
  createSheetsForNestedObjects(jsonData, wb);

  // Save the workbook to an Excel file
  XLSX.writeFile(wb, excelFileName);
}

// Function to flatten nested objects
function flattenObject(obj, prefix = '') {
  if (obj === null || typeof obj !== 'object') {
    return { [prefix]: obj };
  }

  return Object.keys(obj).reduce((acc, key) => {
    const propName = prefix ? `${prefix}.${key}` : key;
    if (typeof obj[key] === 'object' && obj[key] !== null) {
      return { ...acc, ...flattenObject(obj[key], propName) };
    } else {
      return { ...acc, [propName]: obj[key] };
    }
  }, {});
}

// Function to create sheets for nested objects
function createSheetsForNestedObjects(jsonData, wb) {
  for (const key in jsonData) {
    if (jsonData.hasOwnProperty(key) && typeof jsonData[key] === 'object' && jsonData[key] !== null) {
      const nestedData = jsonData[key];

      // Flatten the nested object
      const flattenedData = flattenObject(nestedData);

      // Convert the flattened data to a worksheet
      const wsNested = XLSX.utils.json_to_sheet([flattenedData], { header: Object.keys(flattenedData) });

      // Append the worksheet to the workbook
      XLSX.utils.book_append_sheet(wb, wsNested, key);
    }
  }
}

// Read JSON data from a file(enter path name according to your wish)
const jsonFilePath = './j2e/sample.json';

try {
  const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));

  // Output Excel file name
  const excelFileName = 'output_nested_sheets.xlsx';

  // Convert and save to Excel file
  convertNestedJsonToExcel(jsonData, excelFileName);
  console.log('Conversion completed.');
} catch (error) {
  console.error('Error reading or parsing JSON file:', error.message);
}
