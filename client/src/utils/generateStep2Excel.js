import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import dayjs from 'dayjs';

export function generateStep2Excel({
  workbook,
  selectedSheets,
  // idKeys는 step1에서 선택한 기준 컬럼(실제 결과물의 컬럼명)으로 전달됨
  idKeys = ['지원자번호', '지원직무', '이름'],
  groupSets,      // { 시트명: [ [col1, col2], [col3, col4], ... ] }
  columnOrders,   // { 시트명: [ '자격증명', '발급기관', ... ] }
  sortRules       // { 시트명: { key: '취득시기', method: 'desc'|'asc'|'alpha' } }
}) {
  const newWb = XLSX.utils.book_new();

  selectedSheets.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    // step1에서 전달된 기준 컬럼명을 그대로 baseline으로 사용
    const baselineKeys = idKeys.filter(k => k in data[0]);
    const allCols = Object.keys(data[0]);
    const extraCols = allCols.filter(c => !baselineKeys.includes(c));

    // 그룹 세트: extraCols 중에서 문자+숫자 패턴을 가진 열들을 그룹핑 (숫자 부분 기준)
    const setMap = {};
    extraCols.forEach(col => {
      const match = col.match(/(\D+)(\d+)$/);
      if (match) {
        const key = match[2];
        if (!setMap[key]) setMap[key] = [];
        setMap[key].push(col);
      }
    });

    const flattened = [];

    data.forEach(row => {
      const common = {};
      baselineKeys.forEach(key => {
        common[key] = row[key];
      });
      let added = false;

      Object.keys(setMap).forEach(setId => {
        const set = {};
        setMap[setId].forEach(col => {
          const base = col.replace(/[0-9]+$/, '');
          set[base] = row[col];
        });

        // groupSets가 전달된 경우, 활성 그룹 세트의 첫번째 열을 기준으로 값이 있는지 확인
        const primary = groupSets[sheetName]?.[0]?.[0]?.replace(/[0-9]+$/, '');
        if (primary && !set[primary]) return;

        flattened.push({ ...common, ...set });
        added = true;
      });

      // 그룹 세트에 속하지 않는 경우, columnOrders에 정의된 열들에 대해 '기재 사항 없음'을 넣음
      // 단, 기준 컬럼 및 연번은 제외
      if (!added) {
        const blank = {};
        columnOrders[sheetName].forEach(col => {
          if (!baselineKeys.includes(col) && col !== '연번') {
            blank[col] = '기재 사항 없음';
          }
        });
        flattened.push({ ...common, ...blank });
      }
    });

    // 동일 지원자(기준 컬럼의 첫번째 값) 내에서 연번 부여
    const groupById = {};
    flattened.forEach(row => {
      const key = row[baselineKeys[0]]; // 첫번째 기준 컬럼 사용
      if (!groupById[key]) groupById[key] = [];
      groupById[key].push(row);
    });

    const finalRows = [];

    Object.values(groupById).forEach(rows => {
      const rule = sortRules[sheetName];
      if (rule?.key) {
        rows.sort((a, b) => {
          const aVal = a[rule.key];
          const bVal = b[rule.key];
          if (rule.method === 'alpha') return String(aVal).localeCompare(String(bVal));
          if (!aVal || aVal === '기재 사항 없음') return 1;
          if (!bVal || bVal === '기재 사항 없음') return -1;
          const aDate = dayjs(aVal);
          const bDate = dayjs(bVal);
          if (!aDate.isValid() || !bDate.isValid()) return 0;
          return rule.method === 'asc' ? aDate - bDate : bDate - aDate;
        });
      }

      rows.forEach((row, idx) => {
        finalRows.push({
          ...row,
          연번: idx + 1
        });
      });
    });

    // 기준 컬럼 바로 뒤에 '연번' 컬럼이 오도록 header 순서 지정
    const allKeysFlattened = Object.keys(finalRows[0] || {});
    const otherKeys = allKeysFlattened.filter(key => !baselineKeys.includes(key) && key !== '연번');
    const headerOrder = [...baselineKeys, '연번', ...otherKeys];

    const finalSheet = XLSX.utils.json_to_sheet(finalRows, { header: headerOrder });
    XLSX.utils.book_append_sheet(newWb, finalSheet, sheetName);
  });

  const wbout = XLSX.write(newWb, { bookType: 'xlsx', type: 'array' });
  saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'step2.xlsx');
}
