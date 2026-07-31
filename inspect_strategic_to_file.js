const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('선행수익성 자료/선행점검양식 0722 취합 2.xlsx', { sheets: ['전략제품'], cellStyles: true, cellFormulas: true, cellDates: true, cellNF: true });
const sh = wb.Sheets['전략제품'];

let output = '=== 전략제품 Sheet D to O cols inspection ===\n';

for (let r = 0; r < 70; r++) {
    let rowStr = `Row ${r}: `;
    let hasVal = false;
    for (let c = 0; c < 50; c++) { // Col A ~ Col AX
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = sh[cellRef];
        if (cell && cell.v !== undefined) {
            hasVal = true;
            rowStr += `${cellRef}(Col ${XLSX.utils.encode_col(c)}): v=${cell.v} | w="${cell.w || ''}" | f=${cell.f || ''}  ||  `;
        }
    }
    if (hasVal) {
        output += rowStr + '\n';
    }
}

fs.writeFileSync('strategic_sheet_structure.txt', output, 'utf8');
console.log('Saved strategic_sheet_structure.txt, size:', output.length);
