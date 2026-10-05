const fs = require('fs');
const XLSX = require('xlsx');

const inputFile = 'begrippen.csv';
const outputFile = 'begrippen.md';

// CSV inlezen
const workbook = XLSX.readFile(inputFile, {
  type: 'file',
  raw: false
});

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const rows = XLSX.utils.sheet_to_json(worksheet, {
  defval: ''
});

let markdown = '';

for (const row of rows) {
  const voorkeursterm = String(row['Voorkeursterm'] || '').trim();
  const definitie = String(row['Definitie'] || '').trim();

  if (!voorkeursterm) {
    continue;
  }

  // Kop
  markdown += `# ${voorkeursterm}\n\n`;

  // Definitie
  if (definitie) {
    markdown += `${definitie}\n\n`;
  }

  // Overige gevulde velden
  const fields = Object.entries(row).filter(([key, value]) => {
    return (
      key !== 'Voorkeursterm' &&
      key !== 'Definitie' &&
      String(value).trim() !== ''
    );
  });

  if (fields.length > 0) {
    markdown += '| Veld | Waarde |\n';
    markdown += '|---|---|\n';

    for (const [key, value] of fields) {
      const escapedValue = String(value)
        .replace(/\|/g, '\\|')
        .replace(/\r?\n/g, '<br>');

      markdown += `| ${key} | ${escapedValue} |\n`;
    }

    markdown += '\n';
  }
}

fs.writeFileSync(outputFile, markdown, 'utf8');

console.log(`${outputFile} gegenereerd`);
