import React, { useState, useMemo, useCallback, useRef } from 'react';
import {
  Box, Paper, Typography, Button, Chip, Stack,
  Table, TableHead, TableBody, TableRow, TableCell,
  IconButton, Alert, FormControl, RadioGroup,
  FormControlLabel, Radio, Tooltip,
  Select, MenuItem, InputLabel, Grid
} from '@mui/material';
import AddIcon from '@mui/icons-material/Add';
import RemoveIcon from '@mui/icons-material/Remove';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import dayjs from 'dayjs';

/* ── 유틸 함수들 ────────────────────────── */

function detectSetSize(columnNames) {
  if (columnNames.length <= 1) return null;
  const normalized = columnNames.map(n => String(n).replace(/\d+/g, '#'));
  const N = normalized.length;
  for (let k = 1; k <= Math.floor(N / 2); k++) {
    if (N % k !== 0) continue;
    const first = normalized.slice(0, k);
    let match = true;
    for (let i = k; i < N; i += k) {
      if (normalized.slice(i, i + k).some((v, j) => v !== first[j])) { match = false; break; }
    }
    if (match) return k;
  }
  return null;
}

function autoDetectSetRanges(headers, baseIndices) {
  const baseSet = new Set(baseIndices);
  const remaining = [];
  for (let i = 0; i < headers.length; i++) {
    if (!baseSet.has(i)) remaining.push(i);
  }
  if (remaining.length === 0) return [];

  const contiguous = [];
  let segStart = 0;
  for (let i = 1; i <= remaining.length; i++) {
    if (i === remaining.length || remaining[i] !== remaining[i - 1] + 1) {
      contiguous.push(remaining.slice(segStart, i));
      segStart = i;
    }
  }

  const results = [];
  contiguous.forEach(segment => {
    const names = segment.map(idx => String(headers[idx]).replace(/\d+/g, '#'));
    let pos = 0;
    while (pos < names.length) {
      let bestK = null;
      let bestLen = 0;
      for (let k = 1; k <= Math.floor((names.length - pos) / 2); k++) {
        const pattern = names.slice(pos, pos + k);
        let reps = 1;
        let cursor = pos + k;
        while (cursor + k <= names.length) {
          const next = names.slice(cursor, cursor + k);
          if (next.some((v, j) => v !== pattern[j])) break;
          reps++;
          cursor += k;
        }
        if (reps >= 2 && reps * k > bestLen) {
          bestK = k;
          bestLen = reps * k;
        }
      }
      if (bestK !== null) {
        results.push({
          id: Date.now() + Math.random(),
          start: segment[pos],
          end: segment[pos + bestLen - 1],
          detectedSetSize: bestK,
          manualSetSize: null,
          sortKey: '',
          sortMethod: ''
        });
        pos += bestLen;
      } else {
        pos++;
      }
    }
  });
  return results;
}

function stripSetNumbers(firstSetNames, secondSetNames) {
  if (!secondSetNames || secondSetNames.length === 0) {
    return firstSetNames.map(n => String(n).replace(/\d+$/, '').trim());
  }
  return firstSetNames.map((name, idx) => {
    const s1 = String(name);
    const s2 = String(secondSetNames[idx] ?? '');
    if (!s2) return s1.replace(/\d+$/, '').trim();

    const matches1 = [...s1.matchAll(/\d+/g)];
    const matches2 = [...s2.matchAll(/\d+/g)];
    if (matches1.length !== matches2.length) return s1.replace(/\d+$/, '').trim();

    let result = s1;
    let offset = 0;
    for (let i = 0; i < matches1.length; i++) {
      if (matches1[i][0] !== matches2[i][0]) {
        const start = matches1[i].index + offset;
        const end = start + matches1[i][0].length;
        result = result.slice(0, start) + result.slice(end);
        offset -= matches1[i][0].length;
      }
    }
    return result.trim();
  });
}

function deriveGroupName(headers, range, groupRow) {
  // groupRow가 있으면 (NHR 2행 헤더) 그룹명 직접 사용
  if (groupRow) {
    const g = groupRow[range.start];
    if (g && String(g).trim()) return String(g).trim();
  }
  const rh = headers.slice(range.start, range.end + 1);
  if (rh.length === 0) return '세트';
  let prefix = String(rh[0]);
  for (let i = 1; i < rh.length; i++) {
    while (!String(rh[i]).startsWith(prefix) && prefix.length > 0) prefix = prefix.slice(0, -1);
  }
  prefix = prefix.replace(/[\d\s\-:]+$/, '').trim();
  return prefix || '세트';
}

function rangesOverlap(a, b) {
  return a.start <= b.end && b.start <= a.end;
}

function getEffectiveSetSize(r) {
  return r.manualSetSize ?? r.detectedSetSize;
}

function getSetInfo(r, headers, groupRow) {
  const setSize = getEffectiveSetSize(r);
  if (!setSize) return null;
  const totalCols = r.end - r.start + 1;
  const numSets = Math.floor(totalCols / setSize);
  const first = headers.slice(r.start, r.start + setSize);
  const second = numSets >= 2 ? headers.slice(r.start + setSize, r.start + setSize * 2) : [];
  const cleanNames = stripSetNumbers(first, second);
  return { range: r, setSize, numSets, cleanNames, groupName: deriveGroupName(headers, r, groupRow) };
}

/* ── 컬럼 색상 ── */
const ORANGE_SHADES = ['#FFE0B2', '#FFCC80', '#FFB74D', '#FFA726', '#FF9800'];

function getColumnColor(colIndex, baseIndices, setRanges, rangeFirstClick, selectionMode) {
  if (rangeFirstClick === colIndex) {
    return selectionMode === 'set' ? '#FFA000' : '#64B5F6';
  }
  if (baseIndices.has(colIndex)) return '#BBDEFB';
  for (let i = 0; i < setRanges.length; i++) {
    if (colIndex >= setRanges[i].start && colIndex <= setRanges[i].end) {
      return ORANGE_SHADES[i % ORANGE_SHADES.length];
    }
  }
  return null;
}

/* ══════════════════════════════════════════
   메인 컴포넌트
   ══════════════════════════════════════════ */
export default function VerticalTransformPage() {
  const [headers, setHeaders] = useState([]);
  const [sampleRow, setSampleRow] = useState([]);
  const [allRows, setAllRows] = useState([]);
  const [fileName, setFileName] = useState('');

  // 기본정보: 개별 컬럼 인덱스 배열
  const [baseIndices, setBaseIndices] = useState([]);
  // 2행 헤더 감지 시 그룹행 보존
  const [groupRow, setGroupRow] = useState(null); // string[] | null (0행 그룹명)
  // 세트 범위 (각각 정렬 옵션 포함)
  const [setRanges, setSetRanges] = useState([]);
  // UI 모드
  const [selectionMode, setSelectionMode] = useState(null); // null | 'base' | 'set' | 'fallback'
  const [rangeFirstClick, setRangeFirstClick] = useState(null);
  const [fallbackTargetId, setFallbackTargetId] = useState(null);
  // 다운로드 형식
  const [downloadFormat, setDownloadFormat] = useState('simple');
  const [alertMsg, setAlertMsg] = useState(null);
  const [dragging, setDragging] = useState(false);
  const fileInputRef = useRef(null);

  const baseIndicesSet = useMemo(() => new Set(baseIndices), [baseIndices]);

  /* ── 파일 업로드 ── */
  const processFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name.replace(/\.[^.]+$/, ''));
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const merges = sheet['!merges'] || [];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // 2행 헤더 감지: 병합 풀기 전에 판단해야 함
      // (병합 풀면 0행이 그룹명으로 채워져서 빈칸 비율 체크가 무의미해짐)
      const hasRow0Merges = merges.some(m => m.s.r === 0 && m.e.c > m.s.c);
      const rawRow0 = (raw[0] || []).map(v => (v == null ? '' : String(v)));
      const preEmptyCount = rawRow0.filter(v => !v).length;
      const is2RowHeader = hasRow0Merges ||
        (rawRow0.length > 0 && preEmptyCount / rawRow0.length > 0.5);

      // 병합셀 해제
      merges.forEach(({ s, e }) => {
        const val = raw[s.r]?.[s.c];
        for (let r = s.r; r <= e.r; r++) {
          if (!raw[r]) raw[r] = [];
          for (let c = s.c; c <= e.c; c++) raw[r][c] = val;
        }
      });

      const row0 = (raw[0] || []).map(v => (v == null ? '' : String(v)));
      const row1 = (raw[1] || []).map(v => (v == null ? '' : String(v)));

      if (is2RowHeader) {
        setGroupRow(row0); // 그룹명 (병합 해제 후 채워진 상태)
        setHeaders(row1);  // 실제 컬럼명
        setSampleRow(raw[2] || []);
        setAllRows(raw.slice(2));
      } else {
        setGroupRow(null);
        setHeaders(row0);
        setSampleRow(raw[1] || []);
        setAllRows(raw.slice(1));
      }
      setBaseIndices([]);
      setSetRanges([]);
      setSelectionMode(null);
      setRangeFirstClick(null);
      setAlertMsg(null);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFileInput = useCallback((e) => {
    processFile(e.target.files?.[0]);
  }, [processFile]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && /\.(xlsx|xls|csv)$/i.test(file.name)) processFile(file);
  }, [processFile]);

  const handleReset = useCallback(() => {
    setHeaders([]); setSampleRow([]); setAllRows([]); setFileName('');
    setBaseIndices([]); setSetRanges([]);
    setSelectionMode(null); setRangeFirstClick(null); setAlertMsg(null);
  }, []);

  /* ── 컬럼 클릭 ── */
  const handleColumnClick = useCallback((colIndex) => {
    if (!selectionMode) return;

    // 기본정보: 토글
    if (selectionMode === 'base') {
      setBaseIndices(prev =>
        prev.includes(colIndex) ? prev.filter(i => i !== colIndex) : [...prev, colIndex].sort((a, b) => a - b)
      );
      return;
    }

    // 폴백: 1세트 끝
    if (selectionMode === 'fallback' && fallbackTargetId != null) {
      setSetRanges(prev => prev.map(r => {
        if (r.id !== fallbackTargetId) return r;
        if (colIndex < r.start || colIndex > r.end) return r;
        return { ...r, manualSetSize: colIndex - r.start + 1 };
      }));
      setSelectionMode(null);
      setFallbackTargetId(null);
      return;
    }

    // 세트 범위: 클릭 2번
    if (selectionMode === 'set') {
      if (rangeFirstClick === null) {
        setRangeFirstClick(colIndex);
      } else {
        const start = Math.min(rangeFirstClick, colIndex);
        const end = Math.max(rangeFirstClick, colIndex);
        const nr = { start, end };

        // 겹침 검사
        if (baseIndices.some(i => i >= start && i <= end)) {
          setAlertMsg('세트 범위에 기본정보 컬럼이 포함되어 있습니다.');
          setRangeFirstClick(null);
          return;
        }
        if (setRanges.some(r => rangesOverlap(nr, r))) {
          setAlertMsg('세트 범위가 다른 세트와 겹칩니다.');
          setRangeFirstClick(null);
          return;
        }

        const cols = headers.slice(start, end + 1);
        setSetRanges(prev => [...prev, {
          id: Date.now(),
          start, end,
          detectedSetSize: detectSetSize(cols),
          manualSetSize: null,
          sortKey: '',
          sortMethod: ''
        }]);
        setSelectionMode(null);
        setRangeFirstClick(null);
        setAlertMsg(null);
      }
    }
  }, [selectionMode, rangeFirstClick, baseIndices, setRanges, headers, fallbackTargetId]);

  /* ── 세트 자동 감지 ── */
  const handleAutoDetect = useCallback(() => {
    if (baseIndices.length === 0) return;
    const detected = autoDetectSetRanges(headers, baseIndices);
    if (detected.length === 0) {
      setAlertMsg('자동 감지된 세트 범위가 없습니다.');
      return;
    }
    setSetRanges(detected);
    setAlertMsg(null);
    setSelectionMode(null);
    setRangeFirstClick(null);
  }, [baseIndices, headers]);

  /* ── 세트 크기 조정 ── */
  const adjustSetSize = useCallback((id, delta) => {
    setSetRanges(prev => prev.map(r => {
      if (r.id !== id) return r;
      const cur = getEffectiveSetSize(r);
      if (!cur) return r;
      const total = r.end - r.start + 1;
      return { ...r, manualSetSize: Math.max(1, Math.min(total, cur + delta)) };
    }));
  }, []);

  /* ── 세트별 정렬 옵션 변경 ── */
  const updateSetSort = useCallback((id, field, value) => {
    setSetRanges(prev => prev.map(r => r.id === id ? { ...r, [field]: value } : r));
  }, []);

  /* ── 세트 삭제 ── */
  const removeSetRange = useCallback((id) => {
    setSetRanges(prev => prev.filter(r => r.id !== id));
  }, []);

  /* ── 세트 정보 계산 ── */
  const setInfos = useMemo(() => {
    return setRanges.map(r => getSetInfo(r, headers, groupRow)).filter(Boolean);
  }, [setRanges, headers, groupRow]);

  const allSetsValid = setRanges.length > 0 && setRanges.every(r => getEffectiveSetSize(r));

  /* ── 세트별 세로화 함수 ── */
  const verticalizeOneSet = useCallback((info) => {
    const { range: r, setSize, numSets, cleanNames } = info;
    const outHeaders = ['지원자 연번', ...baseIndices.map(i => headers[i]), '연번', ...cleanNames];

    const flattened = [];
    allRows.forEach(row => {
      const baseVals = baseIndices.map(i => row[i] ?? '');
      const personRows = [];

      for (let setIdx = 0; setIdx < numSets; setIdx++) {
        const vals = [];
        let allEmpty = true;
        for (let col = 0; col < setSize; col++) {
          const v = row[r.start + setIdx * setSize + col] ?? '';
          const s = String(v).trim();
          if (s !== '' && s !== '-') allEmpty = false;
          vals.push(v);
        }
        if (!allEmpty) personRows.push(vals);
      }

      if (personRows.length === 0) {
        flattened.push({ base: baseVals, vals: cleanNames.map(() => '기재 사항 없음'), key: baseVals[0] });
      } else {
        personRows.forEach(v => flattened.push({ base: baseVals, vals: v, key: baseVals[0] }));
      }
    });

    // 그룹핑 → 정렬 → 연번 + 지원자 연번
    const grouped = {};
    const groupOrder = [];
    flattened.forEach(item => {
      const k = item.key ?? '';
      if (!grouped[k]) { grouped[k] = []; groupOrder.push(k); }
      grouped[k].push(item);
    });

    const outRows = [];
    groupOrder.forEach((key, personIdx) => {
      const group = grouped[key];
      if (r.sortKey && r.sortMethod) {
        const idx = cleanNames.indexOf(r.sortKey);
        if (idx >= 0) {
          group.sort((a, b) => {
            const av = a.vals[idx], bv = b.vals[idx];
            if (r.sortMethod === 'alpha') return String(av).localeCompare(String(bv));
            if (!av || av === '기재 사항 없음') return 1;
            if (!bv || bv === '기재 사항 없음') return -1;
            const ad = dayjs(av), bd = dayjs(bv);
            if (!ad.isValid() || !bd.isValid()) return 0;
            return r.sortMethod === 'asc' ? ad - bd : bd - ad;
          });
        }
      }
      group.forEach((item, i) => outRows.push([personIdx + 1, ...item.base, i + 1, ...item.vals]));
    });

    return { outHeaders, outRows, info };
  }, [baseIndices, headers, allRows]);

  /* ── 다운로드 ── */
  const handleDownload = useCallback(() => {
    if (setInfos.length === 0) return;
    const wb = XLSX.utils.book_new();

    // 시트명 중복 처리: 같은 이름이 2개 이상이면 전부 번호 부여
    const rawNames = setInfos.map(info => info.groupName.slice(0, 31) || '시트');
    const nameCount = {};
    rawNames.forEach(n => { nameCount[n] = (nameCount[n] || 0) + 1; });
    const nameIdx = {};
    const sheetNames = rawNames.map(n => {
      if (nameCount[n] <= 1) return n;
      nameIdx[n] = (nameIdx[n] || 0) + 1;
      return `${n.slice(0, 28)}${nameIdx[n]}`;
    });

    setInfos.forEach((info, sheetIdx) => {
      const { outHeaders, outRows } = verticalizeOneSet(info);
      const sheetName = sheetNames[sheetIdx];

      if (downloadFormat === 'nhr') {
        // 2행 헤더 + 병합
        const topHeader = [];
        topHeader.push('기본정보'); // 지원자 연번
        baseIndices.forEach(() => topHeader.push('기본정보'));
        topHeader.push('기본정보'); // 연번
        info.cleanNames.forEach(() => topHeader.push(info.groupName));

        const aoa = [topHeader, outHeaders, ...outRows];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!merges'] = [];
        let ms = 0;
        for (let c = 1; c <= topHeader.length; c++) {
          if (c === topHeader.length || topHeader[c] !== topHeader[ms]) {
            if (c - 1 > ms) ws['!merges'].push({ s: { r: 0, c: ms }, e: { r: 0, c: c - 1 } });
            ms = c;
          }
        }
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      } else {
        const aoa = [outHeaders, ...outRows];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
    });

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), `세로화_${fileName || 'output'}.xlsx`);
  }, [setInfos, verticalizeOneSet, baseIndices, downloadFormat, fileName]);

  /* ── 안내 메시지 ── */
  const modeMessage = useMemo(() => {
    if (!selectionMode) return null;
    if (selectionMode === 'base') return '기본정보 컬럼을 클릭하세요 (토글). 완료되면 "선택 완료" 클릭';
    if (selectionMode === 'fallback') return '세트 범위 안에서 첫 번째 세트의 마지막 컬럼을 클릭하세요';
    if (selectionMode === 'set') {
      return rangeFirstClick === null
        ? '세트 범위의 시작 컬럼을 클릭하세요'
        : `끝 컬럼을 클릭하세요 (시작: ${headers[rangeFirstClick]})`;
    }
    return null;
  }, [selectionMode, rangeFirstClick, headers]);

  /* ══════════ 렌더링 ══════════ */
  return (
    <Box>
      {/* 파일 업로드 */}
      <Box
        onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); setDragging(true); }}
        onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setDragging(false); }}
        onDrop={handleDrop}
        sx={{
          border: '2px dashed',
          borderColor: dragging ? 'primary.main' : 'divider',
          borderRadius: 2,
          p: 2,
          mb: 3,
          textAlign: 'center',
          backgroundColor: dragging ? 'action.hover' : 'transparent',
          transition: 'all 0.2s',
        }}
      >
        <Stack direction="row" spacing={2} alignItems="center" justifyContent="center">
          <Button variant="contained" component="label">
            엑셀 업로드
            <input ref={fileInputRef} type="file" hidden accept=".xlsx,.xls,.csv" onChange={handleFileInput} />
          </Button>
          {fileName ? (
            <>
              <Chip label={fileName} variant="outlined" />
              <Button size="small" color="error" onClick={handleReset}>초기화</Button>
            </>
          ) : (
            <Typography variant="caption" color="text.secondary">
              또는 파일을 여기에 드래그
            </Typography>
          )}
        </Stack>
      </Box>

      {headers.length > 0 && (
        <>
          {/* Step 1: 컬럼 선택 */}
          <Paper variant="outlined" sx={{ p: 2, mb: 2, borderRadius: 2 }}>
            <Typography variant="h6" sx={{ mb: 1.5, fontWeight: 700 }}>
              Step 1: 기본정보 & 세트 범위
            </Typography>

            <Stack direction="row" spacing={1} sx={{ mb: 2 }} flexWrap="wrap">
              <Button
                variant={selectionMode === 'base' ? 'contained' : 'outlined'}
                color="primary"
                onClick={() => { setSelectionMode('base'); setRangeFirstClick(null); setAlertMsg(null); }}
                size="small"
              >
                기본정보 컬럼 선택
              </Button>
              {selectionMode === 'base' && (
                <Button size="small" variant="contained" onClick={() => setSelectionMode(null)}>
                  선택 완료
                </Button>
              )}
              <Button
                variant="contained"
                color="success"
                onClick={handleAutoDetect}
                disabled={baseIndices.length === 0}
                size="small"
              >
                세트 자동 감지
              </Button>
              <Button
                variant={selectionMode === 'set' ? 'contained' : 'outlined'}
                color="warning"
                onClick={() => { setSelectionMode('set'); setRangeFirstClick(null); setAlertMsg(null); }}
                disabled={baseIndices.length === 0}
                size="small"
              >
                세트 범위 수동 추가
              </Button>
              {selectionMode && selectionMode !== 'base' && (
                <Button size="small" color="inherit" onClick={() => {
                  setSelectionMode(null); setRangeFirstClick(null); setFallbackTargetId(null);
                }}>
                  취소
                </Button>
              )}
            </Stack>

            {modeMessage && <Alert severity="info" sx={{ mb: 2 }}>{modeMessage}</Alert>}
            {alertMsg && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setAlertMsg(null)}>{alertMsg}</Alert>}

            {/* 가로 스크롤 테이블 */}
            <Box sx={{ overflowX: 'auto', border: '1px solid', borderColor: 'divider', borderRadius: 1, mb: 2 }}>
              <Table size="small" sx={{ minWidth: Math.max(headers.length * 130, 600) }}>
                <TableHead>
                  <TableRow>
                    {headers.map((h, i) => {
                      const bg = getColumnColor(i, baseIndicesSet, setRanges, rangeFirstClick, selectionMode);
                      return (
                        <TableCell
                          key={i}
                          onClick={() => handleColumnClick(i)}
                          sx={{
                            cursor: selectionMode ? 'pointer' : 'default',
                            backgroundColor: bg || 'inherit',
                            whiteSpace: 'nowrap', fontWeight: 600, fontSize: 12,
                            borderRight: '1px solid', borderRightColor: 'divider',
                            userSelect: 'none',
                            '&:hover': selectionMode ? { opacity: 0.7, transition: '0.15s' } : {}
                          }}
                        >
                          <Tooltip title={`열 ${i + 1}`} arrow placement="top">
                            <span>{h || `(열${i + 1})`}</span>
                          </Tooltip>
                        </TableCell>
                      );
                    })}
                  </TableRow>
                </TableHead>
                <TableBody>
                  <TableRow>
                    {headers.map((_, i) => {
                      const bg = getColumnColor(i, baseIndicesSet, setRanges, rangeFirstClick, selectionMode);
                      return (
                        <TableCell
                          key={i}
                          onClick={() => handleColumnClick(i)}
                          sx={{
                            cursor: selectionMode ? 'pointer' : 'default',
                            backgroundColor: bg || 'inherit',
                            whiteSpace: 'nowrap', fontSize: 12,
                            borderRight: '1px solid', borderRightColor: 'divider',
                            userSelect: 'none', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis'
                          }}
                        >
                          {sampleRow[i] ?? ''}
                        </TableCell>
                      );
                    })}
                  </TableRow>
                </TableBody>
              </Table>
            </Box>

            {/* 선택 현황 */}
            <Stack spacing={1}>
              {baseIndices.length > 0 && (
                <Stack direction="row" spacing={0.5} alignItems="center" flexWrap="wrap">
                  <Typography variant="caption" sx={{ fontWeight: 700, mr: 1 }}>기본정보:</Typography>
                  {baseIndices.map(i => (
                    <Chip
                      key={i}
                      label={headers[i] || `열${i + 1}`}
                      size="small"
                      color="primary"
                      onDelete={() => setBaseIndices(prev => prev.filter(x => x !== i))}
                    />
                  ))}
                </Stack>
              )}

              {setRanges.map((r, idx) => {
                const ss = getEffectiveSetSize(r);
                const total = r.end - r.start + 1;
                const ns = ss ? Math.floor(total / ss) : null;
                const rem = ss ? total % ss : 0;
                return (
                  <Stack key={r.id} direction="row" spacing={1} alignItems="center" flexWrap="wrap">
                    <Chip
                      label={`세트 ${idx + 1}: ${headers[r.start]} ~ ${headers[r.end]} (${total}개)`}
                      color="warning"
                      onDelete={() => removeSetRange(r.id)}
                    />
                    {ss ? (
                      <>
                        <Chip label={`1세트=${ss}컬럼 × ${ns}세트`} size="small" variant="outlined" color="success" />
                        <IconButton size="small" onClick={() => adjustSetSize(r.id, -1)}><RemoveIcon fontSize="small" /></IconButton>
                        <IconButton size="small" onClick={() => adjustSetSize(r.id, 1)}><AddIcon fontSize="small" /></IconButton>
                        {rem > 0 && <Typography variant="caption" color="warning.main">(나머지 {rem}개 무시)</Typography>}
                      </>
                    ) : (
                      <Button size="small" color="error" variant="outlined" onClick={() => {
                        setSelectionMode('fallback'); setFallbackTargetId(r.id); setRangeFirstClick(null);
                      }}>
                        자동 감지 실패 — 1세트 끝 클릭
                      </Button>
                    )}
                  </Stack>
                );
              })}
            </Stack>
          </Paper>

          {/* Step 2: 세트별 정렬 & 다운로드 */}
          {baseIndices.length > 0 && allSetsValid && setInfos.length > 0 && (
            <Paper variant="outlined" sx={{ p: 2, borderRadius: 2 }}>
              <Typography variant="h6" sx={{ mb: 1.5, fontWeight: 700 }}>
                Step 2: 정렬 설정 & 다운로드
              </Typography>

              <Stack spacing={2} sx={{ mb: 3 }}>
                {setInfos.map((info, idx) => (
                  <Paper key={info.range.id} variant="outlined" sx={{ p: 2, borderRadius: 1 }}>
                    <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1.5 }} flexWrap="wrap">
                      <Chip label={info.groupName} color="warning" size="small" />
                      <Chip label={`${info.setSize}컬럼 × ${info.numSets}세트`} size="small" variant="outlined" />
                      <Typography variant="caption" color="text.secondary">
                        컬럼: {info.cleanNames.join(', ')}
                      </Typography>
                    </Stack>

                    <Grid container spacing={2}>
                      <Grid item xs={12} sm={6}>
                        <FormControl fullWidth size="small">
                          <InputLabel>정렬 기준 컬럼</InputLabel>
                          <Select
                            value={info.range.sortKey}
                            label="정렬 기준 컬럼"
                            onChange={(e) => updateSetSort(info.range.id, 'sortKey', e.target.value)}
                          >
                            <MenuItem value="">정렬 안 함</MenuItem>
                            {info.cleanNames.map(n => <MenuItem key={n} value={n}>{n}</MenuItem>)}
                          </Select>
                        </FormControl>
                      </Grid>
                      <Grid item xs={12} sm={6}>
                        <FormControl fullWidth size="small">
                          <InputLabel>정렬 방식</InputLabel>
                          <Select
                            value={info.range.sortMethod}
                            label="정렬 방식"
                            onChange={(e) => updateSetSort(info.range.id, 'sortMethod', e.target.value)}
                          >
                            <MenuItem value="">정렬 안 함</MenuItem>
                            <MenuItem value="desc">내림차순 (최신순)</MenuItem>
                            <MenuItem value="asc">오름차순 (오래된순)</MenuItem>
                            <MenuItem value="alpha">가나다순</MenuItem>
                          </Select>
                        </FormControl>
                      </Grid>
                    </Grid>
                  </Paper>
                ))}
              </Stack>

              {/* 다운로드 형식 */}
              <FormControl sx={{ mb: 2 }}>
                <RadioGroup row value={downloadFormat} onChange={(e) => setDownloadFormat(e.target.value)}>
                  <FormControlLabel value="simple" control={<Radio size="small" />} label="단순 형식 (1행 헤더)" />
                  <FormControlLabel value="nhr" control={<Radio size="small" />} label="NHR 형식 (2행 헤더, 그룹 병합)" />
                </RadioGroup>
              </FormControl>

              <Button variant="contained" color="success" onClick={handleDownload} size="large">
                세로화 엑셀 다운로드
              </Button>
              <Typography variant="caption" color="text.secondary" sx={{ ml: 2 }}>
                세트별로 시트가 나뉘어 하나의 엑셀 파일로 다운로드됩니다.
              </Typography>
            </Paper>
          )}
        </>
      )}
    </Box>
  );
}
