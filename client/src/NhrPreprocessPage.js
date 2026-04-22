import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  Box, Paper, Typography, Button, Stack, Chip,
  FormControl, InputLabel, Select, MenuItem,
  Table, TableHead, TableBody, TableRow, TableCell,
  TextField, Alert, ToggleButton, ToggleButtonGroup,
  Grid
} from '@mui/material';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { detectDelimiters, parseField } from './utils/fieldParse';

const normalize = (s) => String(s ?? '').replace(/\s+/g, '').toLowerCase();
const matchesKeyword = (header, keywords) => {
  const h = normalize(header);
  return keywords.some((k) => h.includes(normalize(k)));
};
const ID_KEYWORDS = ['지원자번호'];
const FIELD_KEYWORDS = ['지원분야', '1지망'];

// 1-based col → Excel column letters (1 → A, 26 → Z, 27 → AA, ...)
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

// ExcelJS cell.value → 표시용 문자열
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

// 주소("A1") → {col, row} (모두 1-based)
const parseAddr = (addr) => {
  const m = addr.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
};

export default function NhrPreprocessPage() {
  const [headers, setHeaders] = useState([]); // row2 텍스트 배열 (0-indexed)
  const [rows, setRows] = useState([]);        // 데이터행 텍스트 (각 row도 0-indexed 배열)
  const [rowOrigIdx, setRowOrigIdx] = useState([]); // 원본 시트에서의 1-based 행번호
  const [columnCount, setColumnCount] = useState(0);
  const [idCol, setIdCol] = useState('');
  const [fieldCol, setFieldCol] = useState('');
  const [mapping, setMapping] = useState({});
  const [digits, setDigits] = useState(7);
  const [separator, setSeparator] = useState('');
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [parseDelimiter, setParseDelimiter] = useState('');
  const [parseMode, setParseMode] = useState('last'); // 'first' | 'last'

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
      setIdCol('');
      setFieldCol('');
      setMapping({});
      setFileName(file.name.replace(/\.(xlsx|xls|csv)$/i, ''));
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
        next[f] = prev[f] || { code: '', transformed: '' };
      });
      return next;
    });
  }, [uniqueFields]);

  // 구분자 자동 감지
  const detectedDelims = useMemo(() => detectDelimiters(uniqueFields), [uniqueFields]);
  useEffect(() => {
    if (detectedDelims.length > 0 && !parseDelimiter) {
      setParseDelimiter(detectedDelims[0].delimiter);
    }
  }, [detectedDelims, parseDelimiter]);

  const handleMapChange = (field, key, value) => {
    setMapping((prev) => ({
      ...prev,
      [field]: { ...(prev[field] || { code: '', transformed: '' }), [key]: value }
    }));
  };

  const handleBulkGenerateCodes = () => {
    setMapping((prev) => {
      const next = { ...prev };
      uniqueFields.forEach((f, i) => {
        next[f] = {
          ...(next[f] || { code: '', transformed: '' }),
          code: toExcelCol(i + 1)
        };
      });
      return next;
    });
  };

  const handleBulkTransformFields = () => {
    setMapping((prev) => {
      const next = { ...prev };
      uniqueFields.forEach((f) => {
        next[f] = {
          ...(next[f] || { code: '', transformed: '' }),
          transformed: parseField(f, parseDelimiter, parseMode)
        };
      });
      return next;
    });
  };

  const mappingComplete =
    uniqueFields.length > 0 &&
    uniqueFields.every((f) => (mapping[f]?.code || '').trim() !== '');

  // 중복 코드 검출
  const duplicateCodes = useMemo(() => {
    const counts = new Map();
    uniqueFields.forEach((f) => {
      const c = (mapping[f]?.code || '').trim();
      if (!c) return;
      counts.set(c, (counts.get(c) || 0) + 1);
    });
    return Array.from(counts.entries())
      .filter(([, n]) => n > 1)
      .map(([c]) => c);
  }, [uniqueFields, mapping]);

  const sameColumnSelected =
    idCol !== '' && fieldCol !== '' && Number(idCol) === Number(fieldCol);

  const canExport =
    idCol !== '' &&
    fieldCol !== '' &&
    !sameColumnSelected &&
    uniqueFields.length > 0 &&
    mappingComplete &&
    duplicateCodes.length === 0;

  const handleExport = async () => {
    setError('');
    if (!canExport) {
      setError('업로드, 컬럼 선택, 모든 지원분야의 코드 입력이 완료되어야 합니다.');
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
      const idIdx = Number(idCol); // 0-based
      const fIdx = Number(fieldCol);
      const fCol1 = fIdx + 1; // 1-based field column in original

      // 정렬 + 수험번호 생성
      const enriched = rows.map((r, i) => {
        const field = String(r[fIdx] ?? '');
        const map = mapping[field] || { code: '', transformed: '' };
        return {
          origRowNum: rowOrigIdx[i],
          field,
          idNum: r[idIdx],
          code: map.code,
          transformed: map.transformed
        };
      });
      const byField = new Map();
      enriched.forEach((item) => {
        if (!byField.has(item.field)) byField.set(item.field, []);
        byField.get(item.field).push(item);
      });
      byField.forEach((list) => {
        list.sort((a, b) => {
          const na = Number(a.idNum);
          const nb = Number(b.idNum);
          const aNum = !isNaN(na) && a.idNum !== '' && a.idNum !== null;
          const bNum = !isNaN(nb) && b.idNum !== '' && b.idNum !== null;
          if (aNum && bNum) return na - nb;
          return String(a.idNum ?? '').localeCompare(String(b.idNum ?? ''));
        });
        list.forEach((item, idx) => {
          const seq = String(idx + 1).padStart(digits, '0');
          item.examNumber = (item.code || '') + (separator || '') + seq;
        });
      });
      const sortedGroups = Array.from(byField.values()).sort((a, b) => {
        const ca = String(a[0]?.code ?? '');
        const cb = String(b[0]?.code ?? '');
        return ca.localeCompare(cb);
      });
      const outputOrder = [];
      sortedGroups.forEach((list) => list.forEach((item) => outputOrder.push(item)));

      // 새 workbook / worksheet
      const outWb = new ExcelJS.Workbook();
      const outWs = outWb.addWorksheet(ws.name || 'Sheet 1');

      // 컬럼 매핑 (모두 1-based)
      const oldToNew = (oldC1) => (oldC1 <= fCol1 ? oldC1 + 1 : oldC1 + 2);
      const newFieldCol1 = oldToNew(fCol1);
      const newMgmtCol1 = newFieldCol1 + 1;

      // 스타일 복사 헬퍼
      const copyCellFull = (srcCell, tgtCell) => {
        tgtCell.value = srcCell.value;
        if (srcCell.style) tgtCell.style = srcCell.style;
      };
      const copyCellWithValue = (srcCell, tgtCell, overrideValue) => {
        tgtCell.value = overrideValue;
        if (srcCell && srcCell.style) tgtCell.style = srcCell.style;
      };

      const origColCount = ws.columnCount;

      // Row 1 (그룹 헤더)
      const srcR1 = ws.getRow(1);
      for (let c = 1; c <= origColCount; c++) {
        const srcCell = srcR1.getCell(c);
        const newC = oldToNew(c);
        copyCellFull(srcCell, outWs.getRow(1).getCell(newC));
      }
      // 새 col 1 (수험번호 위치, row 1은 빈 값) — 인접 A1 스타일 복사
      copyCellWithValue(srcR1.getCell(1), outWs.getRow(1).getCell(1), '');
      // 새 col newMgmtCol1 (지원분야_관리용 row 1 위치, 빈 값) — 인접 J1 스타일 복사
      copyCellWithValue(srcR1.getCell(fCol1), outWs.getRow(1).getCell(newMgmtCol1), '');

      // Row 2 (서브컬럼 헤더)
      const srcR2 = ws.getRow(2);
      for (let c = 1; c <= origColCount; c++) {
        const srcCell = srcR2.getCell(c);
        const newC = oldToNew(c);
        copyCellFull(srcCell, outWs.getRow(2).getCell(newC));
      }
      copyCellWithValue(srcR2.getCell(1), outWs.getRow(2).getCell(1), '수험번호');
      copyCellWithValue(
        srcR2.getCell(fCol1),
        outWs.getRow(2).getCell(newMgmtCol1),
        '지원분야_관리용'
      );

      // 데이터 행 (정렬된 순서로)
      outputOrder.forEach((item, outIdx) => {
        const tgtRowNum = 3 + outIdx;
        const srcRow = ws.getRow(item.origRowNum);
        for (let c = 1; c <= origColCount; c++) {
          const srcCell = srcRow.getCell(c);
          const newC = oldToNew(c);
          copyCellFull(srcCell, outWs.getRow(tgtRowNum).getCell(newC));
        }
        copyCellWithValue(
          srcRow.getCell(1),
          outWs.getRow(tgtRowNum).getCell(1),
          item.examNumber
        );
        copyCellWithValue(
          srcRow.getCell(fCol1),
          outWs.getRow(tgtRowNum).getCell(newMgmtCol1),
          item.transformed || ''
        );
      });

      // Merges 재계산
      const origMergesList = ws.model.merges || [];
      const newMergeRanges = new Set();
      origMergesList.forEach((mergeStr) => {
        const [start, end] = mergeStr.split(':');
        const s = parseAddr(start);
        const ea = parseAddr(end);
        if (!s || !ea) return;
        let newSC = oldToNew(s.col);
        let newEC = oldToNew(ea.col);
        if (s.row === 1 && ea.row === 1) {
          // 기존 row1 머지가 col 1에서 시작하면 새 col 1(수험번호)까지 왼쪽 확장
          if (s.col === 1) newSC = 1;
          // 오른쪽 끝이 지원분야 컬럼이면 지원분야_관리용까지 오른쪽 확장
          if (ea.col === fCol1) newEC = newEC + 1;
        }
        newMergeRanges.add(
          `${toExcelCol(newSC)}${s.row}:${toExcelCol(newEC)}${ea.row}`
        );
      });
      // 원본 row1에 지원분야 컬럼을 포함하는 머지가 전혀 없으면 K:L 머지 생성
      const fieldCoveredByRow1Merge = origMergesList.some((mergeStr) => {
        const [start, end] = mergeStr.split(':');
        const s = parseAddr(start);
        const ea = parseAddr(end);
        return (
          s && ea && s.row === 1 && ea.row === 1 && s.col <= fCol1 && ea.col >= fCol1
        );
      });
      if (!fieldCoveredByRow1Merge) {
        newMergeRanges.add(
          `${toExcelCol(newFieldCol1)}1:${toExcelCol(newMgmtCol1)}1`
        );
      }
      newMergeRanges.forEach((rng) => {
        try {
          outWs.mergeCells(rng);
        } catch (_) {
          // 이미 머지되어 있거나 범위 이슈 발생 시 무시
        }
      });

      // 컬럼 너비 시프트
      for (let c = 1; c <= origColCount; c++) {
        const srcCol = ws.getColumn(c);
        if (srcCol.width != null) {
          outWs.getColumn(oldToNew(c)).width = srcCol.width;
        }
      }
      const w1 = ws.getColumn(1).width;
      if (w1 != null) outWs.getColumn(1).width = w1;
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

      // Freeze panes 복사
      if (ws.views && ws.views.length) {
        outWs.views = ws.views.map((v) => ({ ...v }));
      }

      const outBuf = await outWb.xlsx.writeBuffer();
      saveAs(
        new Blob([outBuf], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }),
        `${fileName || 'nhr'}_전처리.xlsx`
      );
    } catch (err) {
      setError('엑셀 생성 중 오류가 발생했습니다: ' + err.message);
    } finally {
      setBusy(false);
    }
  };

  const exampleExam = useMemo(() => {
    const firstCode =
      (uniqueFields[0] && mapping[uniqueFields[0]]?.code) || 'A';
    return `${firstCode}${separator}${String(1).padStart(digits, '0')}`;
  }, [uniqueFields, mapping, separator, digits]);

  return (
    <Box>
      <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
        <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
          1) NHR 파일 업로드
        </Typography>
        <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
          1행은 그룹 헤더, 2행은 서브 컬럼 헤더로 처리됩니다. 원본 스타일(테두리·색상·머지)은 그대로 유지됩니다.
        </Typography>
        <Stack direction="row" spacing={1} alignItems="center">
          <Button variant="contained" component="label" disabled={busy}>
            파일 선택
            <input type="file" hidden accept=".xlsx,.xls" onChange={handleUpload} />
          </Button>
          {busy && <Typography variant="caption" color="text.secondary">처리 중...</Typography>}
          {fileName && !busy && <Chip label={fileName} size="small" />}
          {headers.length > 0 && !busy && (
            <Chip
              label={`헤더 ${headerOptions.length}개 · 데이터 ${rows.length}행 · 컬럼 ${columnCount}`}
              size="small"
              color="primary"
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
            2) 역할 컬럼 선택
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1.5 }}>
            키워드(띄어쓰기 무시)로 매칭되는 헤더는 색으로 강조됩니다.
          </Typography>
          <Grid container spacing={2}>
            <Grid item xs={12} md={6}>
              <FormControl fullWidth size="small">
                <InputLabel>지원자번호 컬럼</InputLabel>
                <Select value={idCol} label="지원자번호 컬럼" onChange={(e) => setIdCol(e.target.value)}>
                  {headerOptions.map(({ header, index }) => {
                    const hit = matchesKeyword(header, ID_KEYWORDS);
                    return (
                      <MenuItem
                        key={index}
                        value={index}
                        sx={{
                          backgroundColor: hit ? 'rgba(33,150,243,0.15)' : undefined,
                          fontWeight: hit ? 700 : 400
                        }}
                      >
                        [{toExcelCol(index + 1)}] {header}
                      </MenuItem>
                    );
                  })}
                </Select>
              </FormControl>
            </Grid>
            <Grid item xs={12} md={6}>
              <FormControl fullWidth size="small">
                <InputLabel>지원분야 컬럼</InputLabel>
                <Select value={fieldCol} label="지원분야 컬럼" onChange={(e) => setFieldCol(e.target.value)}>
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
            </Grid>
          </Grid>
          {sameColumnSelected && (
            <Alert severity="warning" sx={{ mt: 1.5 }}>
              지원자번호 컬럼과 지원분야 컬럼을 서로 다른 열로 선택해 주세요.
            </Alert>
          )}
        </Paper>
      )}

      {uniqueFields.length > 0 && (
        <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
          <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
            3) 지원분야별 코드 / 변환값 매핑
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
            각 지원분야에 대해 코드(필수)와 변환 후 지원분야(선택)를 입력하세요.
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
              <Button size="small" variant="outlined" onClick={handleBulkGenerateCodes}>
                코드 A~Z 일괄 생성
              </Button>
              <Button
                size="small"
                variant="outlined"
                onClick={handleBulkTransformFields}
                disabled={!parseDelimiter}
              >
                변환 지원분야 일괄 변환
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
                <TableCell sx={{ fontWeight: 700 }}>코드</TableCell>
                <TableCell sx={{ fontWeight: 700 }}>변환 지원분야</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {uniqueFields.map((f) => {
                const code = (mapping[f]?.code || '').trim();
                const isDup = code !== '' && duplicateCodes.includes(code);
                return (
                  <TableRow key={f}>
                    <TableCell>{f}</TableCell>
                    <TableCell sx={{ width: 140 }}>
                      <TextField
                        size="small"
                        value={mapping[f]?.code || ''}
                        onChange={(e) => handleMapChange(f, 'code', e.target.value)}
                        placeholder="예: A"
                        fullWidth
                        error={isDup}
                        helperText={isDup ? '중복 코드' : ''}
                      />
                    </TableCell>
                    <TableCell>
                      <TextField
                        size="small"
                        fullWidth
                        value={mapping[f]?.transformed || ''}
                        onChange={(e) => handleMapChange(f, 'transformed', e.target.value)}
                        placeholder="변환 후 표시될 지원분야"
                      />
                    </TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
          {duplicateCodes.length > 0 && (
            <Alert severity="warning" sx={{ mt: 1.5 }}>
              중복된 코드가 있습니다: {duplicateCodes.join(', ')}. 각 지원분야별로 고유한 코드를 입력해주세요.
            </Alert>
          )}
        </Paper>
      )}

      {uniqueFields.length > 0 && (
        <Paper elevation={1} sx={{ p: 2, mb: 2, borderRadius: 2 }}>
          <Typography variant="subtitle1" sx={{ fontWeight: 700, mb: 1 }}>
            4) 수험번호_숫자 자릿수 / 결합 구분자
          </Typography>
          <Stack direction="row" spacing={3} alignItems="center" flexWrap="wrap">
            <Box>
              <Typography variant="caption" sx={{ display: 'block', mb: 0.5 }}>자릿수</Typography>
              <ToggleButtonGroup size="small" value={digits} exclusive onChange={(_, v) => v && setDigits(v)}>
                <ToggleButton value={6}>6자리</ToggleButton>
                <ToggleButton value={7}>7자리 (기본)</ToggleButton>
              </ToggleButtonGroup>
            </Box>
            <Box>
              <Typography variant="caption" sx={{ display: 'block', mb: 0.5 }}>코드 ↔ 숫자 구분자</Typography>
              <ToggleButtonGroup size="small" value={separator} exclusive onChange={(_, v) => v !== null && setSeparator(v)}>
                <ToggleButton value="">없음 (기본)</ToggleButton>
                <ToggleButton value="-">-</ToggleButton>
                <ToggleButton value="_">_</ToggleButton>
              </ToggleButtonGroup>
            </Box>
            <Box sx={{ ml: 'auto' }}>
              <Typography variant="caption" color="text.secondary">
                예시 수험번호: <b>{exampleExam}</b>
              </Typography>
            </Box>
          </Stack>
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
          {!canExport && (
            <Typography variant="caption" color="text.secondary" sx={{ mr: 'auto' }}>
              컬럼 선택과 모든 지원분야의 코드 입력이 완료되면 다운로드할 수 있습니다.
            </Typography>
          )}
          <Button
            variant="contained"
            color="primary"
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
