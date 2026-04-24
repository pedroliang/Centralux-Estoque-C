
const fs = require('fs');
const path = require('path');

const sharedStringsPath = 'C:/Users/FS/Desktop/antigravity/Centralux Estoque C/xlsx_temp/xl/sharedStrings.xml';
const sheet1Path = 'C:/Users/FS/Desktop/antigravity/Centralux Estoque C/xlsx_temp/xl/worksheets/sheet1.xml';

function parseSharedStrings(xml) {
  const strings = [];
  const matches = xml.matchAll(/<t.*?>(.*?)<\/t>/g);
  for (const match of matches) {
    strings.push(match[1]);
  }
  return strings;
}

function parseSheet(xml, strings) {
  const rows = [];
  const rowMatches = xml.matchAll(/<row r="(\d+)"[^>]*>(.*?)<\/row>/g);
  for (const rowMatch of rowMatches) {
    const rowNum = rowMatch[1];
    const cellsXml = rowMatch[2];
    const cellMatches = cellsXml.matchAll(/<c r="([A-Z]+)(\d+)"(?: t="s")?>(?:<v>(.*?)<\/v>)?/g);
    const row = {};
    for (const cellMatch of cellMatches) {
      const col = cellMatch[1];
      const type = cellMatch[0].includes('t="s"');
      const val = cellMatch[3];
      if (val !== undefined) {
        row[col] = type ? strings[parseInt(val)] : val;
      }
    }
    if (Object.keys(row).length > 0) {
      rows.push(row);
    }
  }
  return rows;
}

try {
  const stringsXml = fs.readFileSync(sharedStringsPath, 'utf8');
  const sheetXml = fs.readFileSync(sheet1Path, 'utf8');
  const strings = parseSharedStrings(stringsXml);
  const data = parseSheet(sheetXml, strings);
  console.log(JSON.stringify(data, null, 2));
} catch (e) {
  console.error(e);
}
