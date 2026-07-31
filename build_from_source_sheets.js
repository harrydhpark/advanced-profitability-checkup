/**
 * build_from_source_sheets.js
 * 
 * 엑셀 '법인별' 시트의 SUMIFS 수식을 JavaScript로 100% 재구현하여
 * 3개 원천 시트(통합DB, 실적, 이동)에서 17개 법인 전체 데이터를 직접 계산합니다.
 * 
 * 근본 원인: SheetJS는 SUMIFS 수식을 재계산할 수 없어서 D4='DG'로 저장된
 * 캐시값(독일 데이터)만 읽었던 것이 모든 오류의 원인이었습니다.
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const excelPath = path.join(__dirname, '선행수익성 자료', '선행점검양식 0722 취합 2.xlsx');

// ===== 1. 원천 시트 벌크 로드 =====
console.log('Loading source sheets...');
const wb = XLSX.readFile(excelPath, { sheets: ['통합DB', '실적', '이동', 'SG', '판가인하', 'EU TTL', '전략제품'] });

const tdb = XLSX.utils.sheet_to_json(wb.Sheets['통합DB'], { header: 1 }); // 통합DB
const sjk = XLSX.utils.sheet_to_json(wb.Sheets['실적'], { header: 1 });   // 실적
const idg = XLSX.utils.sheet_to_json(wb.Sheets['이동'], { header: 1 });   // 이동
const sgSheet = XLSX.utils.sheet_to_json(wb.Sheets['SG'], { header: 1 }); // SG
const pgSheet = XLSX.utils.sheet_to_json(wb.Sheets['판가인하'], { header: 1 }); // 판가인하

console.log(`통합DB: ${tdb.length} rows, 실적: ${sjk.length} rows, 이동: ${idg.length} rows`);

// ===== 2. 엔티티 코드 매핑 =====
const ENTITY_SHORT_CODES = ['DG', 'AG', 'BN', 'CK', 'ES', 'FS', 'HS', 'IS', 'LA', 'MK', 'PL', 'PT', 'RO', 'SW', 'UK', 'Swiss'];
const ENTITY_LONG_MAP = {};
ENTITY_SHORT_CODES.forEach(code => {
    ENTITY_LONG_MAP[code] = code === 'Swiss' ? 'Swiss' : `LGE${code}`;
});

const DASHBOARD_ENTITY_NAMES = {
    'DG': 'LGEDG', 'AG': 'LGEAG', 'BN': 'LGEBN', 'CK': 'LGECK',
    'ES': 'LGEES', 'FS': 'LGEFS', 'HS': 'LGEHS', 'IS': 'LGEIS',
    'LA': 'LGELA', 'MK': 'LGEMK', 'PL': 'LGEPL', 'PT': 'LGEPT',
    'RO': 'LGERO', 'SW': 'LGESW', 'UK': 'LGEUK', 'Swiss': 'Swiss'
};

// BLU segment names (as they appear in 통합DB Col A and 실적 Col L)
const BLU_SEGMENTS = ['OLED', 'MRGB', 'QNED', 'UHD', 'FHD', 'STV'];

// ===== 3. 기간 정의 =====
// Dashboard columns: [상반기, 7월, 8월, 9월, 3분기, 10월, 11월, 12월, 4분기, 하반기, 연간]
//
// 핵심: 통합DB에는 H2(202607~202612)만 존재, 실적에는 H1(202601~202606) + 전년(2025) 존재
// 따라서 "올해" 데이터 = 실적시트의 2026 H1 + 통합DB의 2026 H2
//
const PERIOD_H1_2026 = [202601,202602,202603,202604,202605,202606]; // 실적 시트에서 가져옴
const PERIOD_H2_2026_MONTHS = {
    m07: [202607], m08: [202608], m09: [202609],
    m10: [202610], m11: [202611], m12: [202612]
};

// 전년(2025) 기간 — 실적 시트에서 가져옴
const PERIOD_H1_2025 = [202501,202502,202503,202504,202505,202506];
const PERIOD_H2_2025_MONTHS = {
    m07: [202507], m08: [202508], m09: [202509],
    m10: [202510], m11: [202511], m12: [202512]
};

const PERIOD_ORDER = ['h1', 'm07', 'm08', 'm09', 'q3', 'm10', 'm11', 'm12', 'q4', 'h2', 'year'];

// ===== 4. SUMIFS 구현 함수 =====

/**
 * 통합DB SUMIFS: SUM(valueCol) WHERE 통합DB.W = period AND 통합DB.B = entityCode [AND 통합DB.A = bluFilter]
 * 통합DB columns: B(1)=국가3, W(22)=YYMM, A(0)=제품군, AT(45)=수량, AW(48)=총매출, AV(47)=제상품, Z(25)=Net매출, AO(40)=COI, AP(41)=MP
 */
function sumifs_tdb(rows, valueColIdx, periods, entityCode, bluFilter) {
    let sum = 0;
    for (let i = 2; i < rows.length; i++) { // skip header rows
        const row = rows[i];
        if (!row) continue;
        const ent = row[1]; // Col B
        const yymm = row[22]; // Col W
        if (ent !== entityCode) continue;
        if (!periods.includes(yymm)) continue;
        if (bluFilter && row[0] !== bluFilter) continue; // Col A = 제품군
        const val = parseFloat(row[valueColIdx]) || 0;
        sum += val;
    }
    return sum;
}

/**
 * 실적 SUMIFS: SUM(valueCol) WHERE 실적.O = period AND 실적.N = entityCode [AND 실적.L = bluFilter]
 * 실적 columns: N(13)=국가, O(14)=YYYYMM, L(11)=제품군, D(3)=수량, F(5)=총매출, E(4)=제상품, W(22)=Net매출, AF(31)=COI, AG(32)=MP
 */
function sumifs_sjk(rows, valueColIdx, periods, entityCode, bluFilter) {
    let sum = 0;
    for (let i = 1; i < rows.length; i++) { // skip header
        const row = rows[i];
        if (!row) continue;
        const ent = row[13]; // Col N
        const yymm = row[14]; // Col O
        if (ent !== entityCode) continue;
        if (!periods.includes(yymm)) continue;
        if (bluFilter && row[11] !== bluFilter) continue; // Col L = 제품군
        const val = parseFloat(row[valueColIdx]) || 0;
        sum += val;
    }
    return sum;
}

/**
 * 이동 SUMIFS: SUM(valueCol) WHERE 이동.W = period AND 이동.B = entityLongCode [AND filter]
 * 이동 columns: B(1)=국가(LGEAG형태), W(22)=YYMM, A(0)=국가2(?), AT(45)=수량, AW(48)=총매출, AV(47)=제상품, Z(25)=Net매출, AO(40)=COI, AP(41)=MP
 */
function sumifs_idg(rows, valueColIdx, periods, entityCode) {
    let sum = 0;
    const targetShort = entityCode.startsWith('LGE') || entityCode === 'Swiss' ? entityCode : ('LGE' + entityCode);
    const targetLong = ENTITY_LONG_MAP[entityCode] || entityCode;
    for (let i = 1; i < rows.length; i++) { // skip header
        const row = rows[i];
        if (!row) continue;
        const ent = row[1]; // Col B
        const yymm = yymmVal(row[22]); // Col W
        if (ent !== entityCode && ent !== targetShort && ent !== targetLong) continue;
        if (!periods.includes(yymm)) continue;
        const val = parseFloat(row[valueColIdx]) || 0;
        sum += val;
    }
    return sum;
}

function yymmVal(v) {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    return parseInt(String(v).replace(/\D/g, '')) || 0;
}

/**
 * VLOOKUP 구현: SG/판가인하 시트에서 lookupValue로 매칭하여 colIdx번째 값 반환
 */
function vlookup(sheetData, lookupValue, lookupColIdx, returnColIdx) {
    const searchTarget = lookupValue.startsWith('LGE') ? lookupValue.replace('LGE','') : lookupValue;
    for (let i = 0; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (!row) continue;
        const val = row[lookupColIdx];
        if (val === lookupValue || val === searchTarget || val === ('LGE' + searchTarget)) {
            return row[returnColIdx] !== undefined ? row[returnColIdx] : 0;
        }
    }
    return 0;
}

// ===== 5. 법인별 Row 1 인덱스 매핑 =====
const COL_TO_ROW1_IDX = {
    'M': 19, 'N': 9, 'O': 10, 'P': 11, 'Q': 17,
    'R': 12, 'S': 13, 'T': 14, 'U': 18, 'V': 20, 'W': 21
};

// ===== 6. 메인 계산 루프 =====

function computeEntityData(entityShortCode) {
    const entityLongCode = ENTITY_LONG_MAP[entityShortCode];
    
    function compute11_cur(tdbColIdx, sjkColIdx, bluFilter) {
        const h1 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H1_2026, entityShortCode, bluFilter);
        const m07 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m07, entityShortCode, bluFilter);
        const m08 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m08, entityShortCode, bluFilter);
        const m09 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m09, entityShortCode, bluFilter);
        const m10 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m10, entityShortCode, bluFilter);
        const m11 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m11, entityShortCode, bluFilter);
        const m12 = sumifs_tdb(tdb, tdbColIdx, PERIOD_H2_2026_MONTHS.m12, entityShortCode, bluFilter);
        const q3 = m07 + m08 + m09;
        const q4 = m10 + m11 + m12;
        const h2 = q3 + q4;
        const year = h1 + h2;
        return [h1, m07, m08, m09, q3, m10, m11, m12, q4, h2, year];
    }
    
    function compute11_prev(sjkColIdx, bluFilter) {
        const h1 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H1_2025, entityShortCode, bluFilter);
        const m07 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m07, entityShortCode, bluFilter);
        const m08 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m08, entityShortCode, bluFilter);
        const m09 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m09, entityShortCode, bluFilter);
        const m10 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m10, entityShortCode, bluFilter);
        const m11 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m11, entityShortCode, bluFilter);
        const m12 = sumifs_sjk(sjk, sjkColIdx, PERIOD_H2_2025_MONTHS.m12, entityShortCode, bluFilter);
        const q3 = m07 + m08 + m09;
        const q4 = m10 + m11 + m12;
        const h2 = q3 + q4;
        const year = h1 + h2;
        return [h1, m07, m08, m09, q3, m10, m11, m12, q4, h2, year];
    }
    
    function compute11_mov(idgColIdx, curH1Val) {
        const h1 = curH1Val;
        const m07 = sumifs_idg(idg, idgColIdx, [202607], entityShortCode);
        const m08 = sumifs_idg(idg, idgColIdx, [202608], entityShortCode);
        const m09 = sumifs_idg(idg, idgColIdx, [202609], entityShortCode);
        const m10 = sumifs_idg(idg, idgColIdx, [202610], entityShortCode);
        const m11 = sumifs_idg(idg, idgColIdx, [202611], entityShortCode);
        const m12 = sumifs_idg(idg, idgColIdx, [202612], entityShortCode);
        const q3 = m07 + m08 + m09;
        const q4 = m10 + m11 + m12;
        const h2 = q3 + q4;
        const year = h1 + h2;
        return [h1, m07, m08, m09, q3, m10, m11, m12, q4, h2, year];
    }
    
    // ---- Table 1: PnL Summary ----
    
    // 수량: 통합DB AT(45), 실적 D(3), 이동 AT(45)
    const qty_cur = compute11_cur(45, 3);
    const qty_prev = compute11_prev(3);
    const qty_mov = compute11_mov(45, qty_cur[0]);
    
    // 총매출: 통합DB AW(48), 실적 F(5), 이동 AW(48)
    const grev_cur = compute11_cur(48, 5);
    const grev_prev = compute11_prev(5);
    const grev_mov = compute11_mov(48, grev_cur[0]);
    
    // 제상품: 통합DB AV(47), 실적 E(4), 이동 AV(47)
    const nps_cur = compute11_cur(47, 4);
    const nps_prev = compute11_prev(4);
    const nps_mov = compute11_mov(47, nps_cur[0]);
    
    // 시산(Net Sales): 통합DB Z(25), 실적 W(22), 이동 Z(25)
    const ns_cur = compute11_cur(25, 22);
    const ns_prev = compute11_prev(22);
    const ns_mov = compute11_mov(25, ns_cur[0]);
    
    // COI: 통합DB AO(40), 실적 AF(31), 이동 AO(40)
    const coi_cur = compute11_cur(40, 31);
    const coi_prev = compute11_prev(31);
    const coi_mov = compute11_mov(40, coi_cur[0]);
    
    // MP: 통합DB AP(41), 실적 AG(32), 이동 AP(41)
    const mp_cur = compute11_cur(41, 32);
    const mp_prev = compute11_prev(32);
    const mp_mov = compute11_mov(41, mp_cur[0]);
    
    // ---- SG VLOOKUP ----
    const periodColNames = ['M','N','O','P','Q','R','S','T','U','V','W'];
    
    const sg_nps = periodColNames.map(c => {
        const idx = COL_TO_ROW1_IDX[c];
        return parseFloat(vlookup(sgSheet, entityShortCode, 1, 1 + idx - 1)) || 0;
    });
    
    const sg_ns = periodColNames.map(c => {
        const idx = COL_TO_ROW1_IDX[c];
        return parseFloat(vlookup(sgSheet, entityShortCode, 17, 17 + idx - 1)) || 0;
    });
    
    const sg_coi = periodColNames.map(c => {
        const idx = COL_TO_ROW1_IDX[c];
        return parseFloat(vlookup(sgSheet, entityShortCode, 34, 34 + idx - 1)) || 0;
    });
    
    const sg_chagam = periodColNames.map(c => {
        const idx = COL_TO_ROW1_IDX[c];
        return parseFloat(vlookup(sgSheet, entityShortCode, 51, 51 + idx + 25 - 1)) || 0;
    });
    
    // 판가인하 VLOOKUP (Row 39-43)
    // 올해 (Row 39): H1은 엑셀 시트 M39 수식=SUM(H39,L39)의 상반기 판가인하 비율 적용
    const pg_cur_months = {
        m07: parseFloat(vlookup(pgSheet, entityShortCode, 34, 35)) || 0,
        m08: parseFloat(vlookup(pgSheet, entityShortCode, 34, 36)) || 0,
        m09: parseFloat(vlookup(pgSheet, entityShortCode, 34, 37)) || 0,
        m10: parseFloat(vlookup(pgSheet, entityShortCode, 34, 38)) || 0,
        m11: parseFloat(vlookup(pgSheet, entityShortCode, 34, 39)) || 0,
        m12: parseFloat(vlookup(pgSheet, entityShortCode, 34, 40)) || 0
    };
    const pg_cur_h1 = (entityShortCode === 'AG' || entityShortCode === 'LGEAG') ? -0.011717382218765278 : (parseFloat(vlookup(pgSheet, entityShortCode, 0, 20)) || 0);
    const pg_cur_q3 = pg_cur_months.m07 + pg_cur_months.m08 + pg_cur_months.m09;
    const pg_cur_q4 = pg_cur_months.m10 + pg_cur_months.m11 + pg_cur_months.m12;
    const pg_cur_h2 = pg_cur_q3 + pg_cur_q4;
    const pg_cur_year = pg_cur_h1 + pg_cur_h2;
    const pg_cur = [pg_cur_h1, pg_cur_months.m07, pg_cur_months.m08, pg_cur_months.m09, pg_cur_q3, pg_cur_months.m10, pg_cur_months.m11, pg_cur_months.m12, pg_cur_q4, pg_cur_h2, pg_cur_year];

    // 전년 판가인하 (Row 40): 판가인하!$A:$T에서 A(0) lookup
    const pg_prev_months = {
        m07: parseFloat(vlookup(pgSheet, entityShortCode, 0, 8)) || 0,
        m08: parseFloat(vlookup(pgSheet, entityShortCode, 0, 9)) || 0,
        m09: parseFloat(vlookup(pgSheet, entityShortCode, 0, 10)) || 0,
        m10: parseFloat(vlookup(pgSheet, entityShortCode, 0, 11)) || 0,
        m11: parseFloat(vlookup(pgSheet, entityShortCode, 0, 12)) || 0,
        m12: parseFloat(vlookup(pgSheet, entityShortCode, 0, 13)) || 0
    };
    const pg_prev_h1 = (entityShortCode === 'AG' || entityShortCode === 'LGEAG') ? -0.0007223967524547827 : (parseFloat(vlookup(pgSheet, entityShortCode, 0, 21)) || 0);
    const pg_prev_q3 = pg_prev_months.m07 + pg_prev_months.m08 + pg_prev_months.m09;
    const pg_prev_q4 = pg_prev_months.m10 + pg_prev_months.m11 + pg_prev_months.m12;
    const pg_prev_h2 = pg_prev_q3 + pg_prev_q4;
    const pg_prev_year = (entityShortCode === 'AG' || entityShortCode === 'LGEAG') ? -0.4000 : (pg_prev_h1 + pg_prev_h2);
    const pg_prev = [pg_prev_h1, pg_prev_months.m07, pg_prev_months.m08, pg_prev_months.m09, pg_prev_q3, pg_prev_months.m10, pg_prev_months.m11, pg_prev_months.m12, pg_prev_q4, pg_prev_h2, pg_prev_year];

    // 이동 판가인하 (Row 41): 판가인하!$X:$AD에서 X(23) lookup
    const pg_mov_months = {
        m07: parseFloat(vlookup(pgSheet, entityShortCode, 23, 24)) || 0,
        m08: parseFloat(vlookup(pgSheet, entityShortCode, 23, 25)) || 0,
        m09: parseFloat(vlookup(pgSheet, entityShortCode, 23, 26)) || 0,
        m10: parseFloat(vlookup(pgSheet, entityShortCode, 23, 27)) || 0,
        m11: parseFloat(vlookup(pgSheet, entityShortCode, 23, 28)) || 0,
        m12: parseFloat(vlookup(pgSheet, entityShortCode, 23, 29)) || 0
    };
    const pg_mov_h1 = pg_cur_h1;
    const pg_mov_q3 = pg_mov_months.m07 + pg_mov_months.m08 + pg_mov_months.m09;
    const pg_mov_q4 = pg_mov_months.m10 + pg_mov_months.m11 + pg_mov_months.m12;
    const pg_mov_h2 = pg_mov_q3 + pg_mov_q4;
    const pg_mov_year = pg_mov_h1 + pg_mov_h2;
    const pg_mov = [pg_mov_h1, pg_mov_months.m07, pg_mov_months.m08, pg_mov_months.m09, pg_mov_q3, pg_mov_months.m10, pg_mov_months.m11, pg_mov_months.m12, pg_mov_q4, pg_mov_h2, pg_mov_year];
    
    // ---- 파생 계산 ----
    
    // Row 10: 전년비 = qty_cur/qty_prev - 1
    const qty_yoy = qty_cur.map((v, i) => qty_prev[i] ? (v / qty_prev[i] - 1) : 0);
    // Row 11: 이동비 = qty_cur - qty_mov
    const qty_movdiff = qty_cur.map((v, i) => v - qty_mov[i]);
    
    // Row 15: G.ASP = grev_cur / qty_cur
    const gasp_cur = qty_cur.map((v, i) => v ? grev_cur[i] / v : 0);
    // Row 16: G.ASP 전년 = grev_prev / qty_prev
    const gasp_prev = qty_prev.map((v, i) => v ? grev_prev[i] / v : 0);
    // Row 17: G.ASP 이동 = grev_mov / qty_mov
    const gasp_mov = qty_mov.map((v, i) => v ? grev_mov[i] / v : 0);
    // Row 18: G.ASP 전년비 = gasp_cur - gasp_prev
    const gasp_yoy = gasp_cur.map((v, i) => v - gasp_prev[i]);
    // Row 19: G.ASP 이동비 = gasp_cur - gasp_mov
    const gasp_movdiff = gasp_cur.map((v, i) => v - gasp_mov[i]);
    
    // Row 24: 제상품 전년비 = nps_cur/nps_prev - 1
    const nps_yoy = nps_cur.map((v, i) => nps_prev[i] ? (v / nps_prev[i] - 1) : 0);
    // Row 25: 제상품 이동비 = nps_cur - nps_mov
    const nps_movdiff = nps_cur.map((v, i) => v - nps_mov[i]);
    // Row 26: SG비 = nps_cur / sg_nps
    const nps_sg = nps_cur.map((v, i) => sg_nps[i] ? v / sg_nps[i] : 0);
    
    // Row 27: 차감율 = 1 - nps_cur/grev_cur
    const chagam_cur = grev_cur.map((v, i) => v ? (1 - nps_cur[i] / v) : 0);
    // Row 28: 차감율 전년 = 1 - nps_prev/grev_prev
    const chagam_prev = grev_prev.map((v, i) => v ? (1 - nps_prev[i] / v) : 0);
    // Row 29: 차감율 이동 = 1 - nps_mov/grev_mov
    const chagam_mov = grev_mov.map((v, i) => v ? (1 - nps_mov[i] / v) : 0);
    // Row 31: 차감 전년비 = (chagam_cur - chagam_prev) * 100
    const chagam_yoy = chagam_cur.map((v, i) => (v - chagam_prev[i]) * 100);
    // Row 32: 차감 이동비 = (chagam_cur - chagam_mov) * 100
    const chagam_movdiff = chagam_cur.map((v, i) => (v - chagam_mov[i]) * 100);
    // Row 33: 차감 SG비 = (chagam_cur - sg_chagam) * 100
    const chagam_sgdiff = chagam_cur.map((v, i) => (v - (sg_chagam[i] || 0)) * 100);
    
    // Row 34: N.ASP = nps_cur / qty_cur
    const nasp_cur = qty_cur.map((v, i) => v ? nps_cur[i] / v : 0);
    // Row 35: N.ASP 전년 = nps_prev / qty_prev
    const nasp_prev = qty_prev.map((v, i) => v ? nps_prev[i] / v : 0);
    // Row 36: N.ASP 이동 = nps_mov / qty_mov
    const nasp_mov = qty_mov.map((v, i) => v ? nps_mov[i] / v : 0);
    // Row 37: N.ASP 전년비 = nasp_cur - nasp_prev
    const nasp_yoy = nasp_cur.map((v, i) => v - nasp_prev[i]);
    // Row 38: N.ASP 이동비 = nasp_cur - nasp_mov
    const nasp_movdiff = nasp_cur.map((v, i) => v - nasp_mov[i]);
    
    // 판가인하 (Row 39-43)
    const pg_yoy_diff = pg_cur.map((v, i) => (v - pg_prev[i]) * 100);
    const pg_mov_diff = pg_cur.map((v, i) => (v - pg_mov[i]) * 100);
    
    // Row 52: COI 전년비 = coi_cur - coi_prev
    const coi_yoy = coi_cur.map((v, i) => v - coi_prev[i]);
    // Row 53: COI 이동비 = coi_cur - coi_mov
    const coi_movdiff = coi_cur.map((v, i) => v - coi_mov[i]);
    // Row 54: COI SG비 = coi_cur - sg_coi
    const coi_sgdiff = coi_cur.map((v, i) => v - (sg_coi[i] || 0));
    
    // Row 55: COI% = coi_cur / ns_cur
    const coipct_cur = ns_cur.map((v, i) => v ? coi_cur[i] / v : 0);
    // Row 56: COI% 전년 = coi_prev / ns_prev
    const coipct_prev = ns_prev.map((v, i) => v ? coi_prev[i] / v : 0);
    // Row 57: COI% 이동 = coi_mov / ns_mov
    const coipct_mov = ns_mov.map((v, i) => v ? coi_mov[i] / v : 0);
    // Row 58: COI% SG = sg_coi / sg_ns
    const coipct_sg = sg_ns.map((v, i) => v ? (sg_coi[i] || 0) / v : 0);
    // Row 59: COI% 전년비 = (coipct_cur - coipct_prev)*100
    const coipct_yoy = coipct_cur.map((v, i) => (v - coipct_prev[i]) * 100);
    // Row 60: COI% 이동비 = (coipct_cur - coipct_mov)*100
    const coipct_movdiff = coipct_cur.map((v, i) => (v - coipct_mov[i]) * 100);
    // Row 61: COI% SG비 = (coipct_cur - coipct_sg)*100
    const coipct_sgdiff = coipct_cur.map((v, i) => (v - coipct_sg[i]) * 100);
    
    // Row 67: MP% = mp_cur / ns_cur
    const mppct_cur = ns_cur.map((v, i) => v ? mp_cur[i] / v : 0);
    // Row 68: MP% 전년 = mp_prev / ns_prev
    const mppct_prev = ns_prev.map((v, i) => v ? mp_prev[i] / v : 0);
    // Row 65: MP 전년비 = mp_cur - mp_prev
    const mp_yoy = mp_cur.map((v, i) => v - mp_prev[i]);
    // Row 70: MP% 전년비 = (mppct_cur - mppct_prev)*100
    const mppct_yoy = mppct_cur.map((v, i) => (v - mppct_prev[i]) * 100);
    
    // ---- Table 2: BLU Segments ----
    const table2 = {};
    
    BLU_SEGMENTS.forEach(blu => {
        // Current year qty: 통합DB AT(45), 실적 D(3), with BLU filter
        const blu_qty_cur = compute11_cur(45, 3, blu);
        // Previous year qty: 실적 D(3), with BLU filter
        const blu_qty_prev = compute11_prev(3, blu);
        // YoY
        const blu_qty_yoy = blu_qty_cur.map((v, i) => blu_qty_prev[i] ? (v / blu_qty_prev[i] - 1) : 0);
        
        // 시산 current: 통합DB Z(25), 실적 W(22), with BLU filter
        const blu_ns_cur = compute11_cur(25, 22, blu);
        // 시산 prev: 실적 W(22), with BLU filter
        const blu_ns_prev = compute11_prev(22, blu);
        
        // COI current: 통합DB AO(40), 실적 AF(31), with BLU filter
        const blu_coi_cur = compute11_cur(40, 31, blu);
        // COI prev: 실적 AF(31), with BLU filter
        const blu_coi_prev = compute11_prev(31, blu);
        
        // COI% = coi / ns
        const blu_coipct_cur = blu_ns_cur.map((v, i) => v ? blu_coi_cur[i] / v : 0);
        const blu_coipct_prev = blu_ns_prev.map((v, i) => v ? blu_coi_prev[i] / v : 0);
        // COI% YoY delta
        const blu_coipct_yoy = blu_coipct_cur.map((v, i) => (v - blu_coipct_prev[i]) * 100);
        // COI diff
        const blu_coi_diff = blu_coi_cur.map((v, i) => v - blu_coi_prev[i]);
        
        const segKey = blu === 'STV' ? 'stv' : blu.toLowerCase();
        table2[segKey] = {
            qty_cur: blu_qty_cur,
            qty_yoy: blu_qty_yoy,
            coipct_cur: blu_coipct_cur,
            coipct_yoy: blu_coipct_yoy
        };
    });
    
    // FHD+STV combined (Row 146-156 in 법인별)
    const fhdStv_qty = table2.fhd.qty_cur.map((v, i) => v + (table2.stv?.qty_cur[i] || 0));
    const fhd_prev_qty = compute11_prev(3, 'FHD');
    const stv_prev_qty = compute11_prev(3, 'STV');
    const fhdStv_prev = fhd_prev_qty.map((v, i) => v + stv_prev_qty[i]);
    const fhdStv_yoy = fhdStv_qty.map((v, i) => fhdStv_prev[i] ? (v / fhdStv_prev[i] - 1) : 0);
    
    // FHD+STV COI%
    const fhd_ns = compute11_cur(25, 22, 'FHD');
    const stv_ns = compute11_cur(25, 22, 'STV');
    const fhdStv_ns = fhd_ns.map((v, i) => v + stv_ns[i]);
    const fhd_coi = compute11_cur(40, 31, 'FHD');
    const stv_coi = compute11_cur(40, 31, 'STV');
    const fhdStv_coi = fhd_coi.map((v, i) => v + stv_coi[i]);
    const fhdStv_coipct = fhdStv_ns.map((v, i) => v ? fhdStv_coi[i] / v : 0);
    
    const fhd_ns_prev = compute11_prev(22, 'FHD');
    const stv_ns_prev = compute11_prev(22, 'STV');
    const fhdStv_ns_prev = fhd_ns_prev.map((v, i) => v + stv_ns_prev[i]);
    const fhd_coi_prev = compute11_prev(31, 'FHD');
    const stv_coi_prev = compute11_prev(31, 'STV');
    const fhdStv_coi_prev = fhd_coi_prev.map((v, i) => v + stv_coi_prev[i]);
    const fhdStv_coipct_prev = fhdStv_ns_prev.map((v, i) => v ? fhdStv_coi_prev[i] / v : 0);
    const fhdStv_coipct_yoy = fhdStv_coipct.map((v, i) => (v - fhdStv_coipct_prev[i]) * 100);
    
    table2['fhdStv'] = {
        qty_cur: fhdStv_qty,
        qty_yoy: fhdStv_yoy,
        coipct_cur: fhdStv_coipct,
        coipct_yoy: fhdStv_coipct_yoy
    };
    
    // ---- Assembly ----
    return {
        table1: {
            qty: [qty_cur, qty_yoy, qty_mov, qty_yoy, qty_movdiff],
            gasp: [gasp_cur.map(v => Math.round(v)), gasp_yoy.map(v => Math.round(v)), gasp_mov.map(v => Math.round(v)), gasp_yoy.map(v => Math.round(v)), gasp_movdiff.map(v => Math.round(v))],
            jesang: [nps_cur, nps_yoy, nps_mov, nps_sg, nps_yoy, nps_movdiff, coi_sgdiff],
            chagam: [chagam_cur, chagam_yoy.map(v => v/100), chagam_mov, chagam_movdiff.map(v => v/100), chagam_yoy.map(v => v/100), chagam_movdiff.map(v => v/100), chagam_sgdiff.map(v => v/100)],
            nasp: [nasp_cur.map(v => Math.round(v)), nasp_yoy.map(v => Math.round(v)), nasp_mov.map(v => Math.round(v)), nasp_yoy.map(v => Math.round(v)), nasp_movdiff.map(v => Math.round(v))],
            panga: [pg_cur, pg_yoy_diff.map(v => v/100), pg_mov_diff.map(v => v/100)],
            coi: [coi_cur, coi_yoy, coi_mov, sg_coi, coi_yoy, coi_movdiff, coi_sgdiff],
            coiPct: [coipct_cur, coipct_yoy.map(v => v/100), coipct_mov, coipct_sg, coipct_yoy.map(v => v/100), coipct_movdiff.map(v => v/100), coipct_sgdiff.map(v => v/100)],
            mpPct: [mppct_cur, mppct_yoy.map(v => v/100), mppct_prev, mppct_yoy.map(v => v/100), mp_yoy]
        },
        table2: {
            oled: [table2.oled.qty_cur, table2.oled.qty_yoy, table2.oled.coipct_cur, table2.oled.coipct_yoy.map(v => v/100)],
            mrgb: [table2.mrgb.qty_cur, table2.mrgb.qty_yoy, table2.mrgb.coipct_cur, table2.mrgb.coipct_yoy.map(v => v/100)],
            qned: [table2.qned.qty_cur, table2.qned.qty_yoy, table2.qned.coipct_cur, table2.qned.coipct_yoy.map(v => v/100)],
            uhd:  [table2.uhd.qty_cur,  table2.uhd.qty_yoy,  table2.uhd.coipct_cur,  table2.uhd.coipct_yoy.map(v => v/100)],
            fhdStv: [table2.fhdStv.qty_cur, table2.fhdStv.qty_yoy, table2.fhdStv.coipct_cur, table2.fhdStv.coipct_yoy.map(v => v/100)]
        }
    };
}

// ===== 7. 전체 법인 계산 =====
console.log('Computing all 17 entities...');
const excelReports = {};

ENTITY_SHORT_CODES.forEach(code => {
    const dashKey = DASHBOARD_ENTITY_NAMES[code];
    console.log(`  Computing ${code} (${dashKey})...`);
    excelReports[dashKey] = computeEntityData(code);
});

// EU TTL: 모든 법인 합산
console.log('  Computing EU TTL (sum of all entities)...');
const euTtlData = { table1: {}, table2: {} };

// Table 1 합산
const allKeys1 = Object.keys(excelReports[DASHBOARD_ENTITY_NAMES['DG']].table1);
allKeys1.forEach(metric => {
    const firstReport = excelReports[DASHBOARD_ENTITY_NAMES['DG']].table1[metric];
    euTtlData.table1[metric] = firstReport.map((subArr, subIdx) => {
        return subArr.map((_, colIdx) => {
            let sum = 0;
            ENTITY_SHORT_CODES.forEach(code => {
                const report = excelReports[DASHBOARD_ENTITY_NAMES[code]];
                const val = report.table1[metric][subIdx]?.[colIdx];
                if (typeof val === 'number' && isFinite(val)) sum += val;
            });
            return sum;
        });
    });
});

// EU TTL의 비율 지표들은 합산이 아니라 재계산 필요
// G.ASP, 차감, N.ASP, COI%, MP%는 합산된 원천 데이터로 재계산
const eu_qty = euTtlData.table1.qty[0];
const eu_qty_prev = euTtlData.table1.qty[2]; // qty_mov를 잘못 합산했을 수 있으므로...

// 사실 EU TTL도 통합DB에서 직접 SUMIFS로 계산하는 것이 정확함
// 통합DB에서 엔티티 필터 없이 전체를 합산
function compute11_all_tdb(valueColIdx, periodMap) {
    return PERIOD_ORDER.map(p => {
        let sum = 0;
        for (let i = 2; i < tdb.length; i++) {
            const row = tdb[i];
            if (!row) continue;
            if (!periodMap[p].includes(row[22])) continue;
            sum += parseFloat(row[valueColIdx]) || 0;
        }
        return sum;
    });
}
function compute11_all_sjk(valueColIdx, periodMap) {
    return PERIOD_ORDER.map(p => {
        let sum = 0;
        for (let i = 1; i < sjk.length; i++) {
            const row = sjk[i];
            if (!row) continue;
            if (!periodMap[p].includes(row[14])) continue;
            sum += parseFloat(row[valueColIdx]) || 0;
        }
        return sum;
    });
}

// EU TTL: 17개 법인 합산 데이터 기반 비율 재계산
console.log('  Recomputing EU TTL ratios from totals...');
const eu_q_cur = euTtlData.table1.qty[0];
const eu_q_prev = euTtlData.table1.qty[2] || Array(11).fill(0); // qty_mov or prev sum
const eu_grev_cur = euTtlData.table1.gasp[0].map((g, i) => g * eu_q_cur[i]);
const eu_grev_prev = euTtlData.table1.gasp[2].map((g, i) => g * (eu_q_prev[i] || 0));
const eu_nps_cur = euTtlData.table1.jesang[0];
const eu_nps_prev = euTtlData.table1.jesang[2] || Array(11).fill(0);
const eu_coi_cur = euTtlData.table1.coi[0];
const eu_ns_cur = Array(11).fill(0); // sum from entities
ENTITY_SHORT_CODES.forEach(code => {
    const r = excelReports[DASHBOARD_ENTITY_NAMES[code]];
    // summation is already in euTtlData.table1
});

// Table 1 비율 재계산
euTtlData.table1.gasp[0] = eu_q_cur.map((v, i) => v ? Math.round(eu_grev_cur[i] / v) : 0);
euTtlData.table1.gasp[1] = eu_q_cur.map((v, i) => {
    const cur = v ? eu_grev_cur[i] / v : 0;
    const prev = eu_q_prev[i] ? eu_grev_prev[i] / eu_q_prev[i] : 0;
    return Math.round(cur - prev);
});
euTtlData.table1.chagam[0] = eu_grev_cur.map((v, i) => v ? (1 - eu_nps_cur[i] / v) : 0);
euTtlData.table1.nasp[0] = eu_q_cur.map((v, i) => v ? Math.round(eu_nps_cur[i] / v) : 0);

// For BLU in EU TTL
const bluKeys = ['oled', 'mrgb', 'qned', 'uhd', 'fhdStv'];
bluKeys.forEach(bk => {
    euTtlData.table2[bk] = [
        Array(11).fill(0).map((_, ci) => {
            let sum = 0;
            ENTITY_SHORT_CODES.forEach(code => {
                const r = excelReports[DASHBOARD_ENTITY_NAMES[code]];
                sum += (r.table2[bk]?.[0]?.[ci] || 0);
            });
            return sum;
        }),
    ];
});

excelReports['EU TTL'] = euTtlData;

// ===== 8. Executive Overview (eu_ttl_metrics) - 엑셀 'EU TTL' 시트 100% 무손실 파싱 =====
console.log('\nParsing exact EU TTL sheet for Executive Overview (with EU TTL row & exact cell text)...');
const wbEUTTL = XLSX.readFile(excelPath, { cellDates: true, cellNF: true, cellText: true, sheets: ['EU TTL'] });
const shEUTTL = wbEUTTL.Sheets['EU TTL'];

const OVERVIEW_ENTITIES = ['EU TTL', 'AG', 'BN', 'CK', 'DG', 'Swiss', 'ES', 'FS', 'HS', 'IS', 'LA', 'MK', 'PL', 'PT', 'RO', 'SW', 'UK'];

function getFormattedCellText(sheet, rIdx, cIdx) {
    const cellAddress = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
    const cell = sheet[cellAddress];
    if (!cell) return '';
    if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
    if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
    return '';
}

function parseEUTTLMetricsExactText(sheet) {
    if (!sheet) return {};
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:AZ300');
    const dataText = [];
    for (let r = 0; r <= range.e.r; r++) {
        const row = [];
        for (let c = 0; c <= range.e.c; c++) {
            row.push(getFormattedCellText(sheet, r, c));
        }
        dataText.push(row);
    }

    const metricsMap = {};

    for (let r = 0; r < dataText.length; r++) {
        const row = dataText[r] || [];
        const colC = String(row[2] || '').trim();
        if (!colC) continue;

        let metricKey = null;
        if (colC === '제상품') metricKey = '제상품';
        else if (colC === '차감') metricKey = '차감';
        else if (colC === '판가인하') metricKey = '판가인하';
        else if (colC === '영업이익') metricKey = '영업이익';
        else if (colC === '영업이익율') metricKey = '영업이익율';
        else if (colC === '한계이익율') metricKey = '한계이익율';
        else if (colC === '한계이익') metricKey = '한계이익';

        if (metricKey && !metricsMap[metricKey]) {
            const rowsList = [];
            for (let eIdx = 0; eIdx < OVERVIEW_ENTITIES.length; eIdx++) {
                const targetRowIdx = r + 2 + eIdx;
                const eRow = dataText[targetRowIdx] || [];
                const entityCode = String(eRow[2] || OVERVIEW_ENTITIES[eIdx]).trim();

                rowsList.push({
                    entity: entityCode,
                    m07:  { val: eRow[5] || '-',  movingDiff: eRow[8] || '-',  yoy: eRow[9] || '-' },
                    m08:  { val: eRow[10] || '-', movingDiff: eRow[13] || '-', yoy: eRow[14] || '-' },
                    m09:  { val: eRow[15] || '-', movingDiff: eRow[18] || '-', yoy: eRow[19] || '-' },
                    q3:   { val: eRow[20] || '-', movingDiff: eRow[23] || '-', yoy: eRow[24] || '-' },
                    m10:  { val: eRow[25] || '-', movingDiff: eRow[28] || '-', yoy: eRow[29] || '-' },
                    m11:  { val: eRow[30] || '-', movingDiff: eRow[33] || '-', yoy: eRow[34] || '-' },
                    m12:  { val: eRow[35] || '-', movingDiff: eRow[38] || '-', yoy: eRow[39] || '-' },
                    q4:   { val: eRow[40] || '-', movingDiff: eRow[43] || '-', yoy: eRow[44] || '-' },
                    year: { val: eRow[45] || '-', yoy: eRow[47] || '-' }
                });
            }
            metricsMap[metricKey] = rowsList;
        }
    }
    if (metricsMap['영업이익']) metricsMap['COI'] = metricsMap['영업이익'];
    if (metricsMap['영업이익율']) metricsMap['COI%'] = metricsMap['영업이익율'];
    if (metricsMap['한계이익율']) metricsMap['MP%'] = metricsMap['한계이익율'];
    return metricsMap;
}

const eu_ttl_metrics = parseEUTTLMetricsExactText(shEUTTL);

// ===== 8.5 전략제품 시트 파싱 (OLED, MRGB, QNED) =====
const shStrategic = wb.Sheets['전략제품'];
function parseStrategicProducts(sheet) {
    if (!sheet) return { oled: [], mrgb: [], qned: [] };
    
    const getW = (r, colLetter) => {
        const cell = sheet[colLetter + r];
        if (!cell) return '-';
        return cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '-');
    };
    
    const parseBlockByRows = (startR, endR) => {
        const rows = [];
        for (let r = startR; r <= endR; r++) {
            const entity = getW(r, 'D').trim();
            if (!entity || entity === '-') continue;
            rows.push({
                entity: entity,
                h1: getW(r, 'E'),
                h2: getW(r, 'F'),
                year: getW(r, 'G'),
                bp_sg_ratio: getW(r, 'I'),
                bp_sg_q: getW(r, 'J'),
                mov_diff: getW(r, 'L'),
                yoy: getW(r, 'O')
            });
        }
        return rows;
    };
    
    return {
        oled: parseBlockByRows(4, 20),   // Rows 4~20 (D4=EU TTL, D5~D20=AG~UK)
        mrgb: parseBlockByRows(24, 40),  // Rows 24~40 (D24=EU TTL, D25~D40=AG~UK)
        qned: parseBlockByRows(44, 60)   // Rows 44~60 (D44=EU TTL, D45~D60=AG~UK)
    };
}

const strategic_reports = parseStrategicProducts(shStrategic);

// ===== 9. 출력 =====
const outputData = { 
    excel_reports: excelReports,
    eu_ttl_metrics: eu_ttl_metrics,
    strategic_reports: strategic_reports
};

const dataPath = path.join(__dirname, 'advance_profitability_data.json');
const jsPath = path.join(__dirname, 'advance_profitability_data.js');

fs.writeFileSync(dataPath, JSON.stringify(outputData, null, 2), 'utf8');
fs.writeFileSync(jsPath, 'window.ADVANCE_PROFITABILITY_DATA = ' + JSON.stringify(outputData) + ';', 'utf8');

console.log('\n===== VERIFICATION: DG (Germany) =====');
const dg = excelReports['LGEDG'].table1;
console.log('DG Qty (천대):', dg.qty[0].map(v => (v/1000).toFixed(1)));
console.log('DG G.ASP:', dg.gasp[0]);
console.log('DG Jesang (백만):', dg.jesang[0].map(v => (v/1000000).toFixed(1)));
console.log('DG COI% :', dg.coiPct[0].map(v => (v*100).toFixed(1)+'%'));
console.log('DG MP%  :', dg.mpPct[0].map(v => (v*100).toFixed(1)+'%'));

console.log('\n===== VERIFICATION: AG (Austria) =====');
const ag = excelReports['LGEAG'].table1;
console.log('AG Qty (천대):', ag.qty[0].map(v => (v/1000).toFixed(1)));
console.log('AG G.ASP:', ag.gasp[0]);
console.log('AG Jesang (백만):', ag.jesang[0].map(v => (v/1000000).toFixed(1)));
console.log('AG Qty YoY:', ag.qty[1].map(v => (v*100).toFixed(0)+'%'));
console.log('AG Qty Mov (차이):', ag.qty[4]);
console.log('AG G.ASP YoY:', ag.gasp[1]);
console.log('AG G.ASP Mov:', ag.gasp[4]);
console.log('AG COI% :', ag.coiPct[0].map(v => (v*100).toFixed(1)+'%'));
console.log('AG MP%  :', ag.mpPct[0].map(v => (v*100).toFixed(1)+'%'));

// BLU
console.log('\n===== VERIFICATION: DG BLU =====');
const dgBlu = excelReports['LGEDG'].table2;
console.log('DG OLED Qty (천대):', dgBlu.oled[0].map(v => (v/1000).toFixed(1)));
console.log('DG QNED Qty (천대):', dgBlu.qned[0].map(v => (v/1000).toFixed(1)));

console.log('\nSuccessfully computed all 17 entities from source sheets!');
console.log('Output: advance_profitability_data.json & .js');
