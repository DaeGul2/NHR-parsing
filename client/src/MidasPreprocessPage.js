import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  Box, Paper, Typography, Button, Stack, Chip,
  FormControl, InputLabel, Select, MenuItem,
  Table, TableHead, TableBody, TableRow, TableCell,
  TextField, Alert, ToggleButton, ToggleButtonGroup
} from '@mui/material';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { detectDelimiters, parseField } from './utils/fieldParse';

const normalize = (s) => String(s ?? '').replace(/\s+/g, '').toLowerCase();
const matchesKeyword = (header, keywords) => {
  const h = normalize(header);
  return keywords.some((k) => h.includes(normalize(k)));
};
const FIELD_KEYWORDS = ['지원분야', '1지망'];

const toExcelCol = (n) => {
  let s = '';
  let x = n;
  while (x > 0) {
    x -= 1;
    s = String.fromCharCode(65 + (x % 26)) + s;
    x = Math.floor(x / 26);
  }
  return s;
};

const cellText = (cell) => {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if (v.formula !== undefined) return String(v.result ?? '');
    if (v.text !== undefined) return String(v.text);
    if (v instanceof Date) return v.toISOString();
    if (v.hyperlink) return String(v.text || v.hyperlink);
    return String(v);
  }
  return String(v);
};

const parseAddr = (addr) => {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
};

export default function MidasPreprocessPage() {
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [rowOrigIdx, setRowOrigIdx] = useState([]);
  const [columnCount, setColumnCount] = useState(0);
  const [fieldCol, setFieldCol] = useState('');
  const [mapping, setMapping] = useState({}); // { 원본지원분야: { transformed } }
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [parseDelimiter, setParseDelimiter] = useState('');
  const [parseMode, setParseMode] = useState('last');

  const origWbRef = useRef(null);

  const headerOptions = useMemo(
    () =>
      headers
        .map((h, i) => ({ header: String(h ?? ''), index: i }))
        .filter((item) => item.header.trim() !== ''),
    [headers]
  );

  const handleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    setError('');
    try {
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const ws = wb.worksheets[0];
      const colCount = ws.columnCount;
      const rowCount = ws.rowCount;

      const headerRow = new Array(colCount).fill('');
      const r2 = ws.getRow(2);
      for (let c = 1; c <= colCount; c++) {
        headerRow[c - 1] = cellText(r2.getCell(c));
      }

      const dataRows = [];
      const origIdx = [];
      for (let r = 3; r <= rowCount; r++) {
        const row = ws.getRow(r);
        const vals = new Array(colCount).fill('');
        let hasAny = false;
        for (let c = 1; c <= colCount; c++) {
          const t = cellText(row.getCell(c));
          vals[c - 1] = t;
          if (t !== '') hasAny = true;
        }
        if (hasAny) {
          dataRows.push(vals);
          origIdx.push(r);
        }
      }

      origWbRef.current = wb;
      setColumnCount(colCount);
      setHeaders(headerRow);
      setRows(dataRows);
      setRowOrigIdx(origIdx);
      setFieldCol('');
      setMapping({});
      setParseDelimiter('');
      setFileName(file.name.replace(/\.(xlsx|xls)$/i, ''));
      e.target.value = '';
    } catch (err) {
      setError('파일을 읽는 중 오류가 발생했습니다: ' + err.message);
    } finally {
      setBusy(false);
    }
  };

  const uniqueFields = useMemo(() => {
    if (fieldCol === '' || fieldCol == null) return [];
    const set = new Set();
    rows.forEach((r) => {
      const v = r[fieldCol];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        set.add(String(v));
      }
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows, fieldCol]);

  useEffect(() => {
    setMapping((prev) => {
      const next = {};
      uniqueFields.forEach((f) => {
        next[f] = prev[f] || { transformed: '' };
      });
      return next;
    });
  }, [uniqueFields]);

  const detectedDelims = useMemo(() => detectDelimiters(uniqueFields), [uniqueFields]);
  useEffect(() => {
    if (detectedDelims.length > 0 && !parseDelimiter) {
      setParseDelimiter(detectedDelims[0].delimiter);
    }
  }, [detectedDelims, parseDelimiter]);

  const handleMapChange = (field, value) => {
    setMapping((prev) => ({
      ...prev,
      [field]: { transformed: value }
    }));
  };

  const handleBulkTransformFields = () => {
    setMapping((prev) => {
      const next = { ...prev };
      uniqueFields.forEach((f) => {
        next[f] = { transformed: parseField(f, parseDelimiter, parseMode) };
      });
      return next;
    });
  };

  const canExport = fieldCol !== '' && uniqueFields.length > 0;

  const handleExport = async () => {
    setError('');
    if (!canExport) {
      setError('업로드와 지원분야 컬럼 선택이 필요합니다.');
      return;
    }
    const wb = origWbRef.current;
    if (!wb) {
      setError('원본 파일 정보가 없습니다. 다시 업로드해주세요.');
      return;
    }
    setBusy(true);
    try {
      const ws = wb.worksheets[0];
      const fIdx = Number(fieldCol);
      const fCol1 = fIdx + 1;

      // 출력 순서는 원본 순서 유지 (수험번호·정렬 없음)
      const outputOrder = rows.map((r, i) => ({
        origRowNum: rowOrigIdx[i],
        transformed: mapping[String(r[fIdx] ?? '')]?.transformed || ''
      }));

      const outWb = new ExcelJS.Workbook();
      const outWs = outWb.addWorksheet(ws.name || 'Sheet 1');

      // 컬럼 1개만 삽입: 필드 바로 오른쪽
      const oldToNew = (oldC1) => (oldC1 <= fCol1 ? oldC1 : oldC1 + 1);
      const newMgmtCol1 = fCol1 + 1;

      const copyFull = (srcCell, tgtCell) => {
        tgtCell.value = srcCell.value;
        if (srcCell.style) tgtCell.style = srcCell.style;
      };
      const copyWith = (srcCell, tgtCell, overrideValue) => {
        tgtCell.value = overrideValue;
        if (srcCell && srcCell.style) tgtCell.style = srcCell.style;
      };

      const origColCount = ws.columnCount;

      // Row 1 (그대로 유지, 머지 확장 없음)
      const srcR1 = ws.getRow(1);
      for (let c = 1; c <= origColCount; c++) {
        copyFull(srcR1.getCell(c), outWs.getRow(1).getCell(oldToNew(c)));
      }
      // 새로 삽입된 row1 셀: 인접(필드컬럼) 스타일 복사, 값 비움
      copyWith(srcR1.getCell(fCol1), outWs.getRow(1).getCell(newMgmtCol1), '');

      // Row 2 (헤더)
      const srcR2 = ws.getRow(2);
      for (let c = 1; c <= origColCount; c++) {
        copyFull(srcR2.getCell(c), outWs.getRow(2).getCell(oldToNew(c)));
      }
      copyWith(
        srcR2.getCell(fCol1),
        outWs.getRow(2).getCell(newMgmtCol1),
        '지원분야_관리용'
      );

      // 데이터 (원본 순서 유지)
      outputOrder.forEach((item, outIdx) => {
        const tgtRow = 3 + outIdx;
        const srcRow = ws.getRow(item.origRowNum);
        for (let c = 1; c <= origColCount; c++) {
          copyFull(srcRow.getCell(c), outWs.getRow(tgtRow).getCell(oldToNew(c)));
        }
        copyWith(
          srcRow.getCell(fCol1),
          outWs.getRow(tgtRow).getCell(newMgmtCol1),
          item.transformed || ''
        );
      });

      // 머지: 단순 시프트만 (특별 확장 없음)
      const origMergesList = ws.model.merges || [];
      origMergesList.forEach((mergeStr) => {
        const [start, end] = mergeStr.split(':');
        const s = parseAddr(start);
        const ea = parseAddr(end);
        if (!s || !ea) return;
        const newSC = oldToNew(s.col);
        const newEC = oldToNew(ea.col);
        const rng = `${toExcelCol(newSC)}${s.row}:${toExcelCol(newEC)}${ea.row}`;
        try {
          outWs.mergeCells(rng);
        } catch (_) {}
      });

      // 컬럼 너비
      for (let c = 1; c <= origColCount; c++) {
        const w = ws.getColumn(c).width;
        if (w != null) outWs.getColumn(oldToNew(c)).width = w;
      }
      const wF = ws.getColumn(fCol1).width;
      if (wF != null) outWs.getColumn(newMgmtCol1).width = wF;

      // 행 높이
      for (let r = 1; r <= 2; r++) {
        const h = ws.getRow(r).height;
        if (h != null) outWs.getRow(r).height = h;
      }
      outputOrder.forEach((item, outIdx) => {
        const h = ws.getRow(item.origRowNum).height;
        if (h != null) outWs.getRow(3 + outIdx).height = h;
      });

      if (ws.views && ws.views.length) {
        outWs.views = ws.views.map((v) => ({ ...v }));
      }

      const outBuf = await outWb.xlsx.writeBuffer();
      saveAs(
        new Blob([outBuf], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }),
        `${fileName || 'midas'}_전처리.xlsx`
      );
    } catch (err) {
      setError('엑셀 생성 중 오류가 발생했습니다: ' + err.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Box>
      <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
        <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
          1) 마이다스 파일 업로드
        </Typography>
        <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
          1행은 무시되고 2행을 헤더로 사용합니다. 3행부터 데이터로 처리됩니다.
        </Typography>
        <Stack direction="row" spacing={1} alignItems="center">
          <Button variant="contained" component="label" disabled={busy}>
            파일 선택
            <input type="file" hidden accept=".xlsx" onChange={handleUpload} />
          </Button>
          {busy && <Typography variant="caption" color="text.secondary">처리 중...</Typography>}
          {fileName && !busy && <Chip label={fileName} size="small" />}
          {headers.length > 0 && !busy && (
            <Chip
              label={`헤더 ${headerOptions.length}개 · 데이터 ${rows.length}행 · 컬럼 ${columnCount}`}
              size="small"
              color="secondary"
              variant="outlined"
            />
          )}
        </Stack>
      </Paper>

      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>
          {error}
        </Alert>
      )}

      {headers.length > 0 && (
        <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
          <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
            2) 지원분야 컬럼 선택
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1.5 }}>
            '지원분야'/'1지망' 키워드(띄어쓰기 무시)로 매칭되는 헤더는 색으로 강조됩니다.
          </Typography>
          <FormControl fullWidth size="small">
            <InputLabel>지원분야 컬럼 (예: "지원분야 : 1지망")</InputLabel>
            <Select
              value={fieldCol}
              label='지원분야 컬럼 (예: "지원분야 : 1지망")'
              onChange={(e) => setFieldCol(e.target.value)}
            >
              {headerOptions.map(({ header, index }) => {
                const hit = matchesKeyword(header, FIELD_KEYWORDS);
                return (
                  <MenuItem
                    key={index}
                    value={index}
                    sx={{
                      backgroundColor: hit ? 'rgba(156,39,176,0.15)' : undefined,
                      fontWeight: hit ? 700 : 400
                    }}
                  >
                    [{toExcelCol(index + 1)}] {header}
                  </MenuItem>
                );
              })}
            </Select>
          </FormControl>
        </Paper>
      )}

      {uniqueFields.length > 0 && (
        <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
          <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
            3) 지원분야별 변환값 매핑
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
            각 지원분야에 대해 변환 후 값을 입력하세요. (빈 값이면 빈 문자열로 출력)
          </Typography>

          <Stack spacing={1} sx={{ mb: 1.5 }}>
            <Stack direction="row" spacing={1} alignItems="center" flexWrap="wrap">
              <Typography variant="caption" color="text.secondary">
                감지된 구분자:
              </Typography>
              {detectedDelims.length === 0 && (
                <Typography variant="caption" color="text.disabled">
                  (없음)
                </Typography>
              )}
              <ToggleButtonGroup
                size="small"
                value={parseDelimiter}
                exclusive
                onChange={(_, v) => v !== null && setParseDelimiter(v)}
              >
                {detectedDelims.map((d) => (
                  <ToggleButton key={d.delimiter} value={d.delimiter}>
                    {d.delimiter}&nbsp;
                    <Typography component="span" variant="caption" color="text.secondary">
                      ({d.count}/{uniqueFields.length})
                    </Typography>
                  </ToggleButton>
                ))}
              </ToggleButtonGroup>
              <TextField
                size="small"
                label="직접 입력"
                value={parseDelimiter}
                onChange={(e) => setParseDelimiter(e.target.value)}
                sx={{ width: 120 }}
              />
              <ToggleButtonGroup
                size="small"
                value={parseMode}
                exclusive
                onChange={(_, v) => v && setParseMode(v)}
              >
                <ToggleButton value="first">앞</ToggleButton>
                <ToggleButton value="last">뒤</ToggleButton>
              </ToggleButtonGroup>
            </Stack>
            <Stack direction="row" spacing={1} alignItems="center" flexWrap="wrap">
              <Button
                size="small"
                variant="outlined"
                onClick={handleBulkTransformFields}
                disabled={!parseDelimiter}
              >
                변환값 일괄 채우기
              </Button>
              {uniqueFields[0] && parseDelimiter && (
                <Typography variant="caption" color="text.secondary">
                  예) "{uniqueFields[0]}" → "{parseField(uniqueFields[0], parseDelimiter, parseMode)}"
                </Typography>
              )}
            </Stack>
          </Stack>

          <Table size="small">
            <TableHead>
              <TableRow>
                <TableCell sx={{ fontWeight: 700 }}>원본 지원분야</TableCell>
                <TableCell sx={{ fontWeight: 700 }}>변환 지원분야</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {uniqueFields.map((f) => (
                <TableRow key={f}>
                  <TableCell>{f}</TableCell>
                  <TableCell>
                    <TextField
                      size="small"
                      fullWidth
                      value={mapping[f]?.transformed || ''}
                      onChange={(e) => handleMapChange(f, e.target.value)}
                      placeholder="변환 후 표시될 지원분야"
                    />
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </Paper>
      )}

      {uniqueFields.length > 0 && (
        <Paper
          elevation={1}
          sx={{
            p: 2,
            mb: 2,
            borderRadius: 2,
            display: 'flex',
            justifyContent: 'flex-end',
            gap: 1,
            alignItems: 'center'
          }}
        >
          <Button
            variant="contained"
            color="secondary"
            onClick={handleExport}
            disabled={!canExport || busy}
          >
            {busy ? '생성 중...' : '전처리 엑셀 다운로드'}
          </Button>
        </Paper>
      )}
    </Box>
  );
}
