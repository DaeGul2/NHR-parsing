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

// 공통 서브섹션 스타일
const subSectionSx = {
  p: 1.5,
  mb: 1.5,
  borderRadius: 1.5,
  border: '1px solid',
  borderColor: 'divider',
  backgroundColor: 'action.hover'
};
const exampleBoxSx = {
  mt: 1,
  p: 1,
  borderRadius: 1,
  backgroundColor: 'background.paper',
  border: '1px dashed',
  borderColor: 'divider'
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
  const [seqDigits, setSeqDigits] = useState(2);
  const [seqJoin, setSeqJoin] = useState('_');

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
        next[f] = prev[f] || { sequenceCode: '', transformed: '' };
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

  const handleMapChange = (field, key, value) => {
    setMapping((prev) => ({
      ...prev,
      [field]: {
        ...(prev[field] || { sequenceCode: '', transformed: '' }),
        [key]: value
      }
    }));
  };

  const handleBulkTransformFields = () => {
    setMapping((prev) => {
      const next = { ...prev };
      uniqueFields.forEach((f) => {
        next[f] = {
          ...(next[f] || { sequenceCode: '', transformed: '' }),
          transformed: parseField(f, parseDelimiter, parseMode)
        };
      });
      return next;
    });
  };

  const handleBulkGenerateSequence = () => {
    setMapping((prev) => {
      const next = { ...prev };
      uniqueFields.forEach((f, i) => {
        next[f] = {
          ...(next[f] || { sequenceCode: '', transformed: '' }),
          sequenceCode: String(i + 1).padStart(seqDigits, '0')
        };
      });
      return next;
    });
  };

  // 연번코드 + 조인 + 변환지원분야 (한쪽이 비면 남은 값만)
  const composeMgmtValue = (seq, transformed) => {
    const parts = [seq, transformed]
      .map((p) => String(p ?? '').trim())
      .filter((p) => p !== '');
    return parts.join(seqJoin);
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
      const outputOrder = rows.map((r, i) => {
        const field = String(r[fIdx] ?? '');
        const map = mapping[field] || { sequenceCode: '', transformed: '' };
        return {
          origRowNum: rowOrigIdx[i],
          mgmtValue: composeMgmtValue(map.sequenceCode, map.transformed)
        };
      });

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
          item.mgmtValue || ''
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
            3) 지원분야별 매핑
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 2 }}>
            3개 단계로 연번코드 / 변환 지원분야를 지정하고 최종 지원분야_관리용 값을 확인하세요.
          </Typography>

          {/* ① 변환 지원분야 자동 추출 */}
          <Box sx={subSectionSx}>
            <Typography variant="body2" sx={{ fontWeight: 700, mb: 0.25 }}>
              ① 변환 지원분야 자동 추출
            </Typography>
            <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
              구분자를 선택해 원본 지원분야 문자열을 나눈 뒤, <b>앞/뒤</b> 조각을 변환 지원분야로 일괄 채웁니다.
            </Typography>
            <Grid container spacing={1.5} alignItems="center">
              <Grid item xs={12} md="auto">
                <Typography variant="caption" color="text.secondary" sx={{ mr: 1 }}>
                  감지된 구분자
                </Typography>
                {detectedDelims.length === 0 ? (
                  <Typography component="span" variant="caption" color="text.disabled">
                    (감지되지 않음 — 직접 입력 칸에 입력)
                  </Typography>
                ) : (
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
                )}
              </Grid>
              <Grid item xs={6} md="auto">
                <TextField
                  size="small"
                  label="직접 입력"
                  value={parseDelimiter}
                  onChange={(e) => setParseDelimiter(e.target.value)}
                  sx={{ width: 130 }}
                />
              </Grid>
              <Grid item xs={6} md="auto">
                <ToggleButtonGroup
                  size="small"
                  value={parseMode}
                  exclusive
                  onChange={(_, v) => v && setParseMode(v)}
                >
                  <ToggleButton value="first">앞 조각</ToggleButton>
                  <ToggleButton value="last">뒤 조각</ToggleButton>
                </ToggleButtonGroup>
              </Grid>
              <Grid item xs={12} md="auto">
                <Button
                  size="small"
                  variant="contained"
                  onClick={handleBulkTransformFields}
                  disabled={!parseDelimiter}
                >
                  변환 지원분야 일괄 변환
                </Button>
              </Grid>
            </Grid>
            {uniqueFields[0] && parseDelimiter && (
              <Box sx={exampleBoxSx}>
                <Typography variant="caption" color="text.secondary">
                  예시&nbsp;·&nbsp;원본: <b>"{uniqueFields[0]}"</b>
                  <br />
                  → 구분자 <code>{parseDelimiter}</code> 기준 <b>{parseMode === 'first' ? '앞' : '뒤'} 조각</b>:&nbsp;
                  <b>"{parseField(uniqueFields[0], parseDelimiter, parseMode)}"</b>
                </Typography>
              </Box>
            )}
          </Box>

          {/* ② 연번코드 & 지원분야_관리용 조합 */}
          <Box sx={subSectionSx}>
            <Typography variant="body2" sx={{ fontWeight: 700, mb: 0.25 }}>
              ② 연번코드 & 지원분야_관리용 조합
            </Typography>
            <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
              최종 <b>지원분야_관리용</b> 값은 <code>연번코드 + 조인구분자 + 변환 지원분야</code> 로 조립됩니다.
              한쪽이 비어 있으면 구분자 없이 나머지만 출력됩니다. 연번코드는 숫자가 아니어도 되며 매핑 표에서 직접 수정할 수 있습니다.
            </Typography>
            <Grid container spacing={1.5} alignItems="center">
              <Grid item xs={12} md="auto">
                <Typography variant="caption" color="text.secondary" sx={{ mr: 1 }}>
                  연번 자릿수
                </Typography>
                <ToggleButtonGroup
                  size="small"
                  value={seqDigits}
                  exclusive
                  onChange={(_, v) => v && setSeqDigits(v)}
                >
                  <ToggleButton value={2}>2자리 (기본)</ToggleButton>
                  <ToggleButton value={3}>3자리</ToggleButton>
                </ToggleButtonGroup>
              </Grid>
              <Grid item xs={6} md="auto">
                <TextField
                  size="small"
                  label="조인 구분자"
                  value={seqJoin}
                  onChange={(e) => setSeqJoin(e.target.value)}
                  sx={{ width: 130 }}
                  helperText="기본 '_'"
                />
              </Grid>
              <Grid item xs={6} md="auto">
                <Button size="small" variant="contained" onClick={handleBulkGenerateSequence}>
                  연번코드 일괄 생성
                </Button>
              </Grid>
            </Grid>
            <Box sx={exampleBoxSx}>
              <Typography variant="caption" color="text.secondary">
                예시&nbsp;·&nbsp;연번 <b>{String(1).padStart(seqDigits, '0')}</b>
                &nbsp;+&nbsp;조인 <code>{seqJoin || '(없음)'}</code>
                &nbsp;+&nbsp;변환값 <b>"{uniqueFields[0] && parseDelimiter ? parseField(uniqueFields[0], parseDelimiter, parseMode) : '변환값'}"</b>
                <br />
                → 지원분야_관리용: <b>"{composeMgmtValue(String(1).padStart(seqDigits, '0'), uniqueFields[0] && parseDelimiter ? parseField(uniqueFields[0], parseDelimiter, parseMode) : '변환값')}"</b>
              </Typography>
            </Box>
          </Box>

          {/* ③ 매핑 테이블 */}
          <Typography variant="body2" sx={{ fontWeight: 700, mt: 2, mb: 0.5 }}>
            ③ 매핑 확인 / 직접 수정
          </Typography>
          <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 1 }}>
            위에서 일괄 채운 값을 행 단위로 확인·수정하세요. 마지막 열은 현재 설정 기준 최종 출력값 미리보기입니다.
          </Typography>

          <Table size="small">
            <TableHead>
              <TableRow>
                <TableCell sx={{ fontWeight: 700 }}>원본 지원분야</TableCell>
                <TableCell sx={{ fontWeight: 700 }}>연번코드</TableCell>
                <TableCell sx={{ fontWeight: 700 }}>변환 지원분야</TableCell>
                <TableCell sx={{ fontWeight: 700 }}>최종 지원분야_관리용</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {uniqueFields.map((f) => {
                const seq = mapping[f]?.sequenceCode || '';
                const trans = mapping[f]?.transformed || '';
                const preview = composeMgmtValue(seq, trans);
                return (
                  <TableRow key={f}>
                    <TableCell>{f}</TableCell>
                    <TableCell sx={{ width: 110 }}>
                      <TextField
                        size="small"
                        value={seq}
                        onChange={(e) => handleMapChange(f, 'sequenceCode', e.target.value)}
                        placeholder="예: 01"
                        fullWidth
                      />
                    </TableCell>
                    <TableCell>
                      <TextField
                        size="small"
                        fullWidth
                        value={trans}
                        onChange={(e) => handleMapChange(f, 'transformed', e.target.value)}
                        placeholder="변환 후 표시될 지원분야"
                      />
                    </TableCell>
                    <TableCell sx={{ color: 'text.secondary', fontSize: 12 }}>
                      {preview || <em>(비어 있음)</em>}
                    </TableCell>
                  </TableRow>
                );
              })}
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
