
const XLSX = require('xlsx');
const fs = require('fs');

const filePath = 'C:/Users/FS/Desktop/estoque c loc.xlsx';

try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Process the data to find codes and locations
    // Based on my previous peek, it has columns like CODIGO and LOCALIZACAO
    // It seems to be organized in corridors.
    
    const items = [];
    // The structure might be:
    // A: CORREDOR A | B: (empty?) | D: CORREDOR B | E: (empty?) ...
    // A: CODIGO | B: LOCALIZACAO | D: CODIGO | E: LOCALIZACAO ...
    
    // Let's just output the raw data first to see the structure
    console.log(JSON.stringify(data, null, 2));
} catch (e) {
    console.error(e);
}
