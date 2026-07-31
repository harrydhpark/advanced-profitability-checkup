const XLSX = require('xlsx');
const path = require('path');

const targetExcel = path.join(__dirname, '선행수익성 자료', '선행점검양식 0722 취합 2.xlsx');
const wb = XLSX.readFile(targetExcel, { sheets: ['전략제품'], cellFormulas: true, cellStyles: true, cellDates: true, cellNF: true });
const sh = wb.Sheets['전략제품'];

if (!sh) {
    console.log('Sheet 전략 제품 not found! Available sheets:', wb.SheetNames);
    process.exit(1);
}

const data = XLSX.utils.sheet_to_json(sh, { header: 1, raw: false });

console.log('=== 전략 제품 Sheet Full Inspection ===');
console.log('Total rows:', data.length);
data.forEach((r, idx) => {
    if (r && r.length > 0) {
        const rowItems = [];
        r.forEach((cell, cIdx) => {
            if (cell !== undefined && cell !== null && cell !== '') {
                const cellRef = XLSX.utils.encode_cell({ r: idx, c: cIdx });
                const cellObj = sh[cellRef];
                const wVal = cellObj ? (cellObj.w || cellObj.v) : cell;
                rowItems.push(`${cellRef}(Col ${cIdx}): raw=${cellObj?cellObj.v:''} | w="${wVal}" | f=${cellObj?cellObj.f:''}`);
            }
        });
        if (rowItems.length > 0) {
            console.log(`\nRow ${idx}:`);
            console.log('  ' + rowItems.join('\n  '));
        }
    }
});
