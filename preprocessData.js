import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';

const MEMBERS_DIR = {
  member1: "김한주",
  member2: "김용석",
  member3: "문경수",
  member4: "홍수민"
};

const TARGET_COLUMNS = ['engagement', 'interest', 'excitement', 'stress', 'relaxation'];
const MATCHERS = {
  'engagement': ['engagement', 'eng', '몰입'],
  'interest': ['interest', 'int', '흥미'],
  'excitement': ['excitement', 'exc', '활성'],
  'stress': ['stress', 'str', '스트레스'],
  'relaxation': ['relaxation', 'rel', 'relation', '이완']
};

function normalizeKey(key) {
  return typeof key === 'string' ? key.toLowerCase().trim() : '';
}

function extractNumber(filename) {
  const match = filename.match(/\d+/);
  return match ? parseInt(match[0], 10) : 9999;
}

function processExcelData(filePath) {
  const fileData = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileData, { type: 'buffer' });
  let allRows = [];
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    if (rows && rows.length > 0) {
      allRows = allRows.concat(rows);
    }
  });
  return allRows;
}

function calculateStats(rows, fileName, songNumber) {
  if (!rows || rows.length === 0) return null;
  
  let headerRowIdx = -1;
  let columnIndices = {};
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    let foundCount = 0;
    const tempMap = {};
    
    row.forEach((cell, idx) => {
      const cellStr = normalizeKey(String(cell));
      TARGET_COLUMNS.forEach(target => {
        const matchers = MATCHERS[target];
        if (matchers.some(m => cellStr.includes(m))) {
           tempMap[target] = idx;
           foundCount++;
        }
      });
    });
    
    if (foundCount >= 1) {
      headerRowIdx = i;
      columnIndices = tempMap;
      if (foundCount >= 3) break; 
    }
  }
  
  if (headerRowIdx === -1) return null;
  
  const sums = {};
  const counts = {};
  const rawData = {};
  
  TARGET_COLUMNS.forEach(col => { sums[col] = 0; counts[col] = 0; rawData[col] = []; });
  
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    TARGET_COLUMNS.forEach(col => {
      const colIdx = columnIndices[col];
      if (colIdx !== undefined && row[colIdx] !== null && row[colIdx] !== undefined) {
        let valStr = String(row[colIdx]).replace(/,/g, '');
        const val = parseFloat(valStr);
        if (!isNaN(val)) {
          sums[col] += val;
          counts[col]++;
          rawData[col].push(val);
        }
      }
    });
  }
  
  const stats = { fileName, songNumber, rawData, averages: {} };
  TARGET_COLUMNS.forEach(col => {
    stats.averages[col] = counts[col] > 0 ? (sums[col] / counts[col]) : 0;
  });
  
  return stats;
}

const baseDir = '/Users/a111/Downloads';
const finalData = {};

Object.keys(MEMBERS_DIR).forEach(memberKey => {
  const dirName = MEMBERS_DIR[memberKey];
  const fullPath = path.join(baseDir, dirName);
  
  if (fs.existsSync(fullPath)) {
    console.log(`Processing ${dirName}...`);
    const files = fs.readdirSync(fullPath).filter(f => f.endsWith('.xlsx') || f.endsWith('.csv'));
    files.sort((a, b) => extractNumber(a) - extractNumber(b));
    
    const results = [];
    files.forEach((file, idx) => {
      const filePath = path.join(fullPath, file);
      const rows = processExcelData(filePath);
      const stats = calculateStats(rows, file, idx + 1);
      if (stats) results.push(stats);
    });
    
    finalData[memberKey] = results;
  }
});

fs.writeFileSync('src/preloadedData.json', JSON.stringify(finalData, null, 2));
console.log('Successfully wrote src/preloadedData.json');
