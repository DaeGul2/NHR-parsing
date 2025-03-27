import * as XLSX from 'xlsx';

export function parseExcel(file, callback) {
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const merges = sheet['!merges'] || [];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // 병합 해제
    merges.forEach((merge) => {
      const row = merge.s.r;
      const start = merge.s.c;
      const end = merge.e.c;
      const value = raw[row][start];
      for (let i = start; i <= end; i++) {
        raw[row][i] = value;
      }
    });

    const headerRow = raw[0];
    const rows = raw.slice(1);

    callback({ headerRow, rows });
  };

  reader.readAsArrayBuffer(file);
}
