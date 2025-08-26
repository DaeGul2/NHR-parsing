import React, {
  useMemo,
  useState,
  useRef,
  useEffect,
  useCallback,
  memo,
} from 'react';
import {
  Dialog, AppBar, Toolbar, IconButton, Typography, Button,
  Box, ListItem, Checkbox, TextField, Divider, Paper,
  Chip, Stack, Tooltip
} from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

/** ===============================
 *  가상 스크롤 리스트 (좌측 헤더용)
 *  - items: {i:number, name:string}[]
 *  - itemHeight: px (정수)
 *  - height: 컨테이너 높이
 *  - renderRow: (item, isSelected, onClick) => JSX
 * =============================== */
const VirtualizedList = memo(function VirtualizedList({
  items,
  height = 400,
  itemHeight = 32,
  selectedSet,
  onRowClick,
  lastClickedIdxRef,
  onShiftRangeSelect,
}) {
  const containerRef = useRef(null);
  const [scrollTop, setScrollTop] = useState(0);

  const totalHeight = items.length * itemHeight;
  const startIndex = Math.max(0, Math.floor(scrollTop / itemHeight) - 5);
  const endIndex = Math.min(
    items.length,
    Math.ceil((scrollTop + height) / itemHeight) + 5
  );
  const visible = items.slice(startIndex, endIndex);

  const onScroll = (e) => setScrollTop(e.currentTarget.scrollTop);

  return (
    <Box
      ref={containerRef}
      onScroll={onScroll}
      sx={{
        position: 'relative',
        height,
        overflow: 'auto',
        border: '1px solid #eee',
        borderRadius: 1,
      }}
    >
      <Box sx={{ height: totalHeight, position: 'relative' }}>
        {visible.map((item, idx) => {
          const top = (startIndex + idx) * itemHeight;
          const isSelected = selectedSet.has(item.i);
          return (
            <RowItem
              key={item.i}
              top={top}
              height={itemHeight}
              item={item}
              isSelected={isSelected}
              onRowClick={onRowClick}
              lastClickedIdxRef={lastClickedIdxRef}
              onShiftRangeSelect={onShiftRangeSelect}
            />
          );
        })}
      </Box>
    </Box>
  );
});

const RowItem = memo(function RowItem({
  top,
  height,
  item,
  isSelected,
  onRowClick,
  lastClickedIdxRef,
  onShiftRangeSelect,
}) {
  const handleClick = (e) => {
    if (e.shiftKey && lastClickedIdxRef.current != null) {
      onShiftRangeSelect(lastClickedIdxRef.current, item.i, e);
    } else {
      onRowClick(item.i, e);
    }
    lastClickedIdxRef.current = item.i;
  };

  return (
    <Box
      onClick={handleClick}
      sx={{
        position: 'absolute',
        left: 0,
        right: 0,
        top,
        height,
        display: 'flex',
        alignItems: 'center',
        px: 1,
        gap: 1,
        cursor: 'pointer',
        '&:hover': { backgroundColor: 'rgba(0,0,0,0.03)' },
      }}
    >
      <Checkbox
        size="small"
        checked={isSelected}
        onClick={(e) => e.stopPropagation()}
        onChange={(e) => handleClick(e)}
      />
      <Tooltip title={`열 인덱스: ${item.i + 1}`}>
        <Typography variant="body2" noWrap>
          {item.name || `(빈 헤더) - 열${item.i + 1}`}
        </Typography>
      </Tooltip>
    </Box>
  );
});

/** 개별 하위명 입력 행 (메모) */
const SubNameRow = memo(function SubNameRow({ idx, value, onChange, onDelete, disableDelete }) {
  const handleChange = (e) => onChange(idx, e.target.value);
  const handleDelete = () => onDelete(idx);
  return (
    <Stack direction="row" spacing={1}>
      <TextField
        size="small"
        placeholder={`하위명 ${idx + 1}`}
        value={value}
        onChange={handleChange}
        fullWidth
      />
      <Button variant="outlined" onClick={handleDelete} disabled={disableDelete}>
        삭제
      </Button>
    </Stack>
  );
});

/**
 * NHR 형식 변환 모달
 * - 엑셀 업로드
 * - 좌: 남은(미선택) 헤더 목록 (가상 스크롤)
 * - 우: 그룹/세트로 옮겨진 구성 미리보기
 * - 하위 세트명 CSV(,) 한 번에 입력 + 디바운스
 * - 선택한 컬럼만 내보내며, 최상단에 그룹/세트명이 헤더로 추가, 병합(merge) 처리
 */
export default function NhrTransformModal({ open, onClose }) {
  // 원본
  const [masterHeaders, setMasterHeaders] = useState([]); // 모든 헤더 (원본 첫 행)
  const [dataRows, setDataRows] = useState([]);           // 데이터 영역
  // 가용(미선택) 컬럼 인덱스
  const [availableIdx, setAvailableIdx] = useState([]);
  // 좌측에서 선택 중인 인덱스들
  const [selectedIdx, setSelectedIdx] = useState([]);

  // 그룹들: { name, indices:number[] }
  const [groups, setGroups] = useState([]);
  // 세트들: { baseName, subNames:string[], chunks: Array<{start:number, indices:number[], newHeaders:string[]}> }
  const [sets, setSets] = useState([]);

  // 입력 상태
  const [groupName, setGroupName] = useState('');
  const [setBaseName, setSetBaseName] = useState('');

  // 하위명: 개별 배열 + CSV 텍스트(디바운스)
  const [subNames, setSubNames] = useState(['']);
  const [subNamesCSV, setSubNamesCSV] = useState(''); // 표시/편집용
  const csvDebounceRef = useRef(null);

  // Shift 범위 선택용 참조
  const lastClickedIdxRef = useRef(null);

  // 업로드
  const handleUpload = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const merges = sheet['!merges'] || [];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // 병합 해제하여 첫 행(헤더) 채움
      merges.forEach(({ s, e }) => {
        const row = s.r;
        const val = raw[row][s.c];
        for (let i = s.c; i <= e.c; i++) raw[row][i] = val;
      });

      const headers = (raw[0] || []).map((v) => (v == null ? '' : String(v)));
      const body = raw.slice(1);

      setMasterHeaders(headers);
      setDataRows(body);
      setAvailableIdx(headers.map((_, i) => i));
      setSelectedIdx([]);
      setGroups([]);
      setSets([]);
      setGroupName('');
      setSetBaseName('');
      setSubNames(['']);
      setSubNamesCSV('');
      lastClickedIdxRef.current = null;
    };
    reader.readAsArrayBuffer(file);
  }, []);

  /** 좌측 리스트에 보여줄 헤더들 (가상 스크롤 대상) */
  const leftHeaders = useMemo(
    () => availableIdx.map((i) => ({ i, name: masterHeaders[i] ?? `열${i + 1}` })),
    [availableIdx, masterHeaders]
  );

  /** 선택 상태를 Set으로 (가상 리스트 행에 빠르게 전달) */
  const selectedSet = useMemo(() => new Set(selectedIdx), [selectedIdx]);

  /** CSV 입력 → 디바운스 적용해서 subNames로 반영 */
  useEffect(() => {
    // subNames -> CSV 표시 동기화 (외부 변경 시)
    setSubNamesCSV(subNames.join(', '));
  }, [subNames]);

  const handleSubNamesCSVChange = useCallback((val) => {
    setSubNamesCSV(val);
    if (csvDebounceRef.current) clearTimeout(csvDebounceRef.current);
    csvDebounceRef.current = setTimeout(() => {
      const arr = val
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      setSubNames(arr.length ? arr : ['']);
    }, 200);
  }, []);

  /** 개별 하위명 입력 */
  const changeSubName = useCallback((k, v) => {
    setSubNames((prev) => {
      const next = [...prev];
      next[k] = v;
      return next;
    });
  }, []);

  const addSubName = useCallback(() => {
    setSubNames((prev) => [...prev, '']);
  }, []);

  const removeSubName = useCallback((k) => {
    setSubNames((prev) => (prev.length <= 1 ? prev : prev.filter((_, i) => i !== k)));
  }, []);

  /** 단일 토글 */
  const onRowClick = useCallback((idx, e) => {
    setSelectedIdx((prev) =>
      prev.includes(idx) ? prev.filter((i) => i !== idx) : [...prev, idx]
    );
  }, []);

  /** Shift 범위 선택 */
  const onShiftRangeSelect = useCallback((lastIdx, currentIdx, e) => {
    // 현재 표시 순서는 leftHeaders의 index 순서
    const ordered = leftHeaders.map(({ i }) => i);
    const start = ordered.indexOf(lastIdx);
    const end = ordered.indexOf(currentIdx);
    if (start === -1 || end === -1) return;
    const [a, b] = start <= end ? [start, end] : [end, start];
    const range = ordered.slice(a, b + 1);
    setSelectedIdx((prev) => Array.from(new Set([...prev, ...range])));
  }, [leftHeaders]);

  const selectAll = useCallback(
    () => setSelectedIdx([...availableIdx]),
    [availableIdx]
  );
  const clearSelected = useCallback(
    () => setSelectedIdx([]),
    []
  );

  /** 그룹 만들기 */
  const handleCreateGroup = useCallback(() => {
    const name = groupName.trim();
    if (!name || selectedIdx.length === 0) return;
    const sorted = [...selectedIdx].sort((a, b) => a - b);

    setGroups((prev) => [...prev, { name, indices: sorted }]);

    // 좌측에서 제거
    const remove = new Set(sorted);
    setAvailableIdx((prev) => prev.filter((i) => !remove.has(i)));
    setSelectedIdx([]);
    setGroupName('');
  }, [groupName, selectedIdx]);

  /** 세트 만들기 */
  const handleCreateSet = useCallback(() => {
    const base = setBaseName.trim();
    const cleanSubs = subNames.map((s) => s.trim()).filter(Boolean);
    if (!base || cleanSubs.length === 0 || selectedIdx.length === 0) return;

    const sorted = [...selectedIdx].sort((a, b) => a - b);
    const chunkSize = cleanSubs.length;
    const chunks = [];
    let ok = true;

    for (let i = 0; i < sorted.length; i += chunkSize) {
      const slice = sorted.slice(i, i + chunkSize);
      if (slice.length !== chunkSize) {
        ok = false;
        break;
      }
      const partIndex = i / chunkSize + 1;
      const newHeaders = cleanSubs.map((sub) => `${base} - ${sub}${partIndex}`);
      chunks.push({ start: i, indices: slice, newHeaders });
    }
    if (!ok) return;

    setSets((prev) => [...prev, { baseName: base, subNames: cleanSubs, chunks }]);

    // 좌측에서 제거
    const remove = new Set(sorted);
    setAvailableIdx((prev) => prev.filter((i) => !remove.has(i)));
    setSelectedIdx([]);
    setSetBaseName('');
    // subNames 는 그대로 두고(CSV입력 유지), 필요시 초기화하려면 아래 주석 해제
    // setSubNames(['']);
    // setSubNamesCSV('');
  }, [setBaseName, subNames, selectedIdx]);

  /** 우측 미리보기 (메모) */
  const previewRight = useMemo(() => {
    const items = [];

    // 그룹
    for (let gi = 0; gi < groups.length; gi++) {
      const g = groups[gi];
      items.push({
        type: 'group',
        key: `group-${gi}`,
        title: g.name,
        cols: g.indices.map((i) => masterHeaders[i] ?? `열${i + 1}`),
      });
    }

    // 세트
    for (let si = 0; si < sets.length; si++) {
      const s = sets[si];
      const all = [];
      for (let ci = 0; ci < s.chunks.length; ci++) {
        const ch = s.chunks[ci];
        for (let j = 0; j < ch.indices.length; j++) {
          const colIdx = ch.indices[j];
          const oldName = masterHeaders[colIdx] ?? `열${colIdx + 1}`;
          const newName = ch.newHeaders[j];
          all.push(`${oldName} → ${newName}`);
        }
      }
      items.push({
        type: 'set',
        key: `set-${si}`,
        title: s.baseName,
        cols: all,
      });
    }

    return items;
  }, [groups, sets, masterHeaders]);

  /** 내보내기: 선택된 컬럼만, 두 줄 헤더 + 병합 */
  const handleDownload = useCallback(() => {
    if (masterHeaders.length === 0) return;

    const topHeader = [];
    const bottomHeader = [];
    const colPickers = [];

    // 그룹
    for (let g = 0; g < groups.length; g++) {
      const grp = groups[g];
      for (let k = 0; k < grp.indices.length; k++) {
        const idx = grp.indices[k];
        topHeader.push(grp.name);
        bottomHeader.push(masterHeaders[idx] ?? `열${idx + 1}`);
        colPickers.push({ srcIdx: idx });
      }
    }

    // 세트
    for (let s = 0; s < sets.length; s++) {
      const set = sets[s];
      for (let ci = 0; ci < set.chunks.length; ci++) {
        const ch = set.chunks[ci];
        for (let j = 0; j < ch.indices.length; j++) {
          const idx = ch.indices[j];
          topHeader.push(set.baseName);
          bottomHeader.push(ch.newHeaders[j]);
          colPickers.push({ srcIdx: idx });
        }
      }
    }

    if (colPickers.length === 0) return;

    const out = [];
    out.push(topHeader);
    out.push(bottomHeader);
    for (let r = 0; r < dataRows.length; r++) {
      const row = dataRows[r];
      const newRow = colPickers.map(({ srcIdx }) => row[srcIdx]);
      out.push(newRow);
    }

    const ws = XLSX.utils.aoa_to_sheet(out);

    // 상단 헤더 병합: 같은 값이 연속되는 범위 병합
    ws['!merges'] = ws['!merges'] || [];
    let start = 0;
    for (let c = 1; c <= topHeader.length; c++) {
      if (c === topHeader.length || topHeader[c] !== topHeader[start]) {
        if (c - 1 > start) {
          ws['!merges'].push({ s: { r: 0, c: start }, e: { r: 0, c: c - 1 } });
        }
        start = c;
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'nhr');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'nhr_transformed.xlsx');
  }, [masterHeaders, groups, sets, dataRows]);

  const canMakeGroup = useMemo(
    () => groupName.trim() && selectedIdx.length > 0,
    [groupName, selectedIdx.length]
  );

  const subCount = useMemo(
    () => subNames.map((s) => s.trim()).filter(Boolean).length,
    [subNames]
  );

  const canMakeSet = useMemo(() => {
    return (
      setBaseName.trim() &&
      subCount > 0 &&
      selectedIdx.length > 0 &&
      selectedIdx.length % subCount === 0
    );
  }, [setBaseName, subCount, selectedIdx.length]);

  return (
    <Dialog fullScreen open={open} onClose={onClose}>
      <AppBar sx={{ position: 'relative' }}>
        <Toolbar>
          <IconButton edge="start" color="inherit" onClick={onClose}>
            <CloseIcon />
          </IconButton>
          <Typography sx={{ ml: 2, flex: 1 }} variant="h6">
            NHR 형식 변환
          </Typography>

          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
            <Button
              variant="contained"
              color="secondary"
              onClick={handleDownload}
              disabled={masterHeaders.length === 0}
            >
              선택 컬럼만 다운로드
            </Button>
          </Box>
        </Toolbar>
      </AppBar>

      <Box sx={{ p: 2, display: 'flex', gap: 2, height: 'calc(100% - 64px)' }}>
        {/* 좌측 패널 */}
        <Paper sx={{ width: '40%', p: 2, display: 'flex', flexDirection: 'column', overflow: 'hidden', gap: 1 }}>
          <Typography variant="h6">1) 엑셀 업로드</Typography>
          <Button variant="outlined" component="label" sx={{ mb: 1 }}>
            엑셀 선택
            <input
              type="file"
              hidden
              accept=".xlsx"
              onChange={(e) => e.target.files?.[0] && handleUpload(e.target.files[0])}
            />
          </Button>

          <Divider />

          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mt: 1 }}>
            <Typography variant="h6" sx={{ flex: 1 }}>
              2) 헤더 선택
            </Typography>
            <Button size="small" onClick={selectAll} disabled={leftHeaders.length === 0}>
              전체선택
            </Button>
            <Button size="small" onClick={clearSelected} disabled={selectedIdx.length === 0}>
              선택해제
            </Button>
          </Box>
          <Typography variant="caption" color="text.secondary">
            팁: 첫 항목 클릭 후 <b>Shift</b> 누른 채로 다른 항목을 클릭하면 사이 범위를 전체 선택.
          </Typography>

          {/* 가상 스크롤 리스트 */}
          <VirtualizedList
            items={leftHeaders}
            height={400}
            itemHeight={32}
            selectedSet={selectedSet}
            onRowClick={onRowClick}
            lastClickedIdxRef={lastClickedIdxRef}
            onShiftRangeSelect={onShiftRangeSelect}
          />

          <Divider sx={{ my: 1 }} />

          {/* 그룹 만들기 */}
          <Typography variant="h6">3) 그룹으로 묶기</Typography>
          <Stack direction="row" spacing={1}>
            <TextField
              size="small"
              placeholder="그룹명 입력"
              value={groupName}
              onChange={(e) => setGroupName(e.target.value)}
              fullWidth
            />
            <Button variant="contained" onClick={handleCreateGroup} disabled={!canMakeGroup}>
              그룹 만들기
            </Button>
          </Stack>

          <Divider sx={{ my: 1 }} />

          {/* 세트 만들기 */}
          <Typography variant="h6">4) 세트로 묶기</Typography>
          <TextField
            size="small"
            placeholder="세트명 입력 (예: 경력사항)"
            value={setBaseName}
            onChange={(e) => setSetBaseName(e.target.value)}
            fullWidth
            sx={{ mb: 1 }}
          />

          {/* ✅ CSV 한 줄 입력 + 디바운스 */}
          <TextField
            size="small"
            placeholder="하위 항목명들을 쉼표(,)로 입력: 회사명, 재직기간, 시작일, 종료일, 부서명, 직급, 담당업무"
            value={subNamesCSV}
            onChange={(e) => handleSubNamesCSVChange(e.target.value)}
            fullWidth
            sx={{ mb: 1 }}
          />
          <Typography variant="caption" color="text.secondary" sx={{ mb: 1 }}>
            예: <code>회사명, 재직기간, 재직기간 시작일, 재직기간 종료일, 부서명, 직급, 담당업무</code>
          </Typography>

          {/* 개별 입력(메모 + 최소 렌더) */}
          <Stack spacing={1} sx={{ maxHeight: 140, overflow: 'auto', mb: 1 }}>
            {subNames.map((v, idx) => (
              <SubNameRow
                key={idx}
                idx={idx}
                value={v}
                onChange={changeSubName}
                onDelete={removeSubName}
                disableDelete={subNames.length === 1}
              />
            ))}
          </Stack>

          <Stack direction="row" spacing={1} sx={{ mb: 1 }}>
            <Button variant="outlined" onClick={addSubName}>하위명 추가</Button>
            <Chip
              label={`선택된 열: ${selectedIdx.length}개 / 세트 단위: ${subCount || 0}`}
              size="small"
            />
          </Stack>

          <Button variant="contained" onClick={handleCreateSet} disabled={!canMakeSet}>
            세트 만들기
          </Button>
        </Paper>

        {/* 우측 패널 */}
        <Paper sx={{ flex: 1, p: 2, overflow: 'auto' }}>
          <Typography variant="h6" gutterBottom>구성 미리보기</Typography>

          {/* 그룹 미리보기 */}
          <Box sx={{ mb: 3 }}>
            <Typography variant="subtitle1" gutterBottom>그룹</Typography>
            {groups.length === 0 && (
              <Typography variant="body2" color="text.secondary">아직 생성된 그룹이 없습니다.</Typography>
            )}
            {groups.map((g, gi) => (
              <Box key={gi} sx={{ mb: 1, p: 1, border: '1px solid #eee', borderRadius: 1 }}>
                <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
                  <Chip label="GROUP" size="small" />
                  <Typography variant="subtitle2">{g.name}</Typography>
                  <Chip label={`${g.indices.length} cols`} size="small" />
                </Stack>
                <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
                  {g.indices.map((i) => (
                    <Chip key={i} label={masterHeaders[i] ?? `열${i + 1}`} size="small" />
                  ))}
                </Box>
              </Box>
            ))}
          </Box>

          {/* 세트 미리보기 */}
          <Box>
            <Typography variant="subtitle1" gutterBottom>세트</Typography>
            {sets.length === 0 && (
              <Typography variant="body2" color="text.secondary">아직 생성된 세트가 없습니다.</Typography>
            )}
            {sets.map((s, si) => (
              <Box key={si} sx={{ mb: 2, p: 1, border: '1px solid #eee', borderRadius: 1 }}>
                <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
                  <Chip label="SET" size="small" color="primary" />
                  <Typography variant="subtitle2">{s.baseName}</Typography>
                  <Chip label={`단위: ${s.subNames.length}`} size="small" />
                  <Chip label={`세트 수: ${s.chunks.length}`} size="small" />
                </Stack>
                {s.chunks.map((ch, ci) => (
                  <Box key={ci} sx={{ mb: 1 }}>
                    <Typography variant="body2" sx={{ mb: 0.5 }}>
                      세트 {ci + 1}
                    </Typography>
                    <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
                      {ch.indices.map((colIdx, j) => (
                        <Chip
                          key={`${colIdx}-${j}`}
                          size="small"
                          label={`${masterHeaders[colIdx] ?? `열${colIdx + 1}`} → ${ch.newHeaders[j]}`}
                        />
                      ))}
                    </Box>
                  </Box>
                ))}
              </Box>
            ))}
          </Box>

          <Divider sx={{ my: 2 }} />

          <Typography variant="body2" color="text.secondary">
            다운로드 시 선택된 컬럼만 포함되며, 최상단에는 그룹/세트명이 병합되어 표시되고 바로 아래 행에 상세 컬럼명이 배치됩니다.
          </Typography>
        </Paper>
      </Box>
    </Dialog>
  );
}
