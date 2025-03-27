import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export function generateStep1Excel({ headerRow, rows, selectedGroups, excludeGroups, idCols }) {
  const wb = XLSX.utils.book_new();

  // 그룹별 컬럼 인덱스 맵 만들기
  const groupMap = {};
  headerRow.forEach((group, idx) => {
    if (!group) return;
    if (!groupMap[group]) groupMap[group] = [];
    groupMap[group].push(idx);
  });

  selectedGroups.forEach((group) => {
    const colIndices = groupMap[group];
    if (!colIndices || colIndices.length === 0) return;

    const groupRows = rows.map((row) => {
      const base = [row[idCols['지원자번호']], row[idCols['지원직무']], row[idCols['이름']]];
      const groupData = colIndices.map(i => row[i]);
      return excludeGroups.includes(group) ? groupData : [...base, ...groupData];
    });

    const ws = XLSX.utils.aoa_to_sheet(groupRows);
    XLSX.utils.book_append_sheet(wb, ws, group);
  });

  // 원본도 추가
  const wsRaw = XLSX.utils.aoa_to_sheet([headerRow, ...rows]);
  XLSX.utils.book_append_sheet(wb, wsRaw, 'rawdata');

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'step1.xlsx');
}
