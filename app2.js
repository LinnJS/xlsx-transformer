const fs = require('fs');
const xlsx = require('xlsx');

// workbook is a folder (spreadsheet)
// sheet is a file in a folder (sheet in spreadsheet)

const workbook = xlsx.readFile('VISN-17_Facility_HS_before.xlsx');

let worksheets = {};

for (const sheetName of workbook.SheetNames) {
  console.log('worksheets: ', worksheets);
  console.log('sheetName: ', sheetName);
  //
  worksheets[sheetName] = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName]);
}

console.log('CSV', worksheets.Sheet1);

const newBook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newBook, worksheets.Sheet1, 'Sheet1');
xlsx.writeFile(newBook, 'new-book.csv');
