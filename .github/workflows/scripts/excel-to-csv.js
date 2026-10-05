const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('begrippen.xlsx');

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const csv = XLSX.utils.sheet_to_csv(worksheet, {
  FS: ';'
});

fs.writeFileSync(
  'begrippen.csv',
  '\uFEFF' + csv,
  'utf8'
);

console.log(`Werkblad "${sheetName}" omgezet naar begrippen.csv`);
