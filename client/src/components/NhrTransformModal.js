import React, { useMemo, useState } from 'react';
import {
    Dialog, AppBar, Toolbar, IconButton, Typography, Button,
    Box, List, ListItem, Checkbox, TextField, Divider, Paper,
    Chip, Stack, Tooltip
} from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

/**
 * NHR 형식 변환 모달
 * - 엑셀 업로드
 * - 좌: 남은(미선택) 헤더 목록
 * - 우: 그룹/세트로 옮겨진 구성 미리보기
 * - 선택한 컬럼만 내보내며, 최상단에 그룹/세트명을 헤더로 추가, 병합(merge) 처리
 */
export default function NhrTransformModal({ open, onClose }) {
    // 원본
    const [masterHeaders, setMasterHeaders] = useState([]); // 모든 헤더 (원본 첫 행)
    const [dataRows, setDataRows] = useState([]);           // 데이터 영역
    // 가용(미선택) 컬럼 인덱스
    const [availableIdx, setAvailableIdx] = useState([]);
    // 좌측에서 선택 중인 인덱스들
    const [selectedIdx, setSelectedIdx] = useState([]);
    const [lastClickedIdx, setLastClickedIdx] = useState(null);
    // 그룹들: { name, indices:number[] }
    const [groups, setGroups] = useState([]);
    // 세트들: { baseName, subNames:string[], chunks: Array<{start:number, indices:number[], newHeaders:string[]}> }
    const [sets, setSets] = useState([]);

    // 세트 만들기 UI 입력 상태
    const [groupName, setGroupName] = useState('');
    const [setBaseName, setSetBaseName] = useState('');
    const [subNames, setSubNames] = useState(['']); // 동적 입력

    // 업로드
    const handleUpload = (file) => {
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
        };
        reader.readAsArrayBuffer(file);
    };

   
    const selectAll = () => setSelectedIdx([...availableIdx]);
    const clearSelected = () => setSelectedIdx([]);
    // 수정: toggleSelect가 이벤트를 받도록
    const toggleSelect = (idx, e) => {
        // 현재 화면에 보이는 왼쪽 목록의 순서(연속 범위 계산용)
        const ordered = leftHeaders.map(({ i }) => i);

        if (e?.shiftKey && lastClickedIdx !== null) {
            // 두 점 사이 범위를 모두 선택
            const start = ordered.indexOf(lastClickedIdx);
            const end = ordered.indexOf(idx);
            if (start !== -1 && end !== -1) {
                const [a, b] = start <= end ? [start, end] : [end, start];
                const range = ordered.slice(a, b + 1);
                setSelectedIdx((prev) => Array.from(new Set([...prev, ...range])));
            }
        } else {
            // 단일 토글
            setSelectedIdx((prev) =>
                prev.includes(idx) ? prev.filter((i) => i !== idx) : [...prev, idx]
            );
        }

        setLastClickedIdx(idx);
    };

    // 그룹 만들기
    const handleCreateGroup = () => {
        if (!groupName.trim() || selectedIdx.length === 0) return;
        const sorted = [...selectedIdx].sort((a, b) => a - b);
        setGroups((prev) => [...prev, { name: groupName.trim(), indices: sorted }]);
        // 좌측에서 제거
        const remain = availableIdx.filter((i) => !sorted.includes(i));
        setAvailableIdx(remain);
        setSelectedIdx([]);
        setGroupName('');
    };

    // 하위명 필드 추가/삭제/변경
    const addSubName = () => setSubNames((prev) => [...prev, '']);
    const removeSubName = (k) =>
        setSubNames((prev) => prev.filter((_, i) => i !== k));
    const changeSubName = (k, val) =>
        setSubNames((prev) => prev.map((v, i) => (i === k ? val : v)));

    // 세트 만들기
    const handleCreateSet = () => {
        const base = setBaseName.trim();
        const cleanSubs = subNames.map(s => s.trim()).filter(Boolean);
        if (!base || cleanSubs.length === 0 || selectedIdx.length === 0) return;

        const sorted = [...selectedIdx].sort((a, b) => a - b);
        const chunkSize = cleanSubs.length;
        const chunks = [];
        let ok = true;

        for (let i = 0; i < sorted.length; i += chunkSize) {
            const slice = sorted.slice(i, i + chunkSize);
            if (slice.length !== chunkSize) {
                ok = false; break;
            }
            const partIndex = i / chunkSize + 1;
            const newHeaders = cleanSubs.map((sub) => `${base} - ${sub}${partIndex}`);
            chunks.push({ start: i, indices: slice, newHeaders });
        }
        if (!ok) return; // 정확히 n개 단위로 떨어지지 않으면 생성하지 않음

        setSets((prev) => [...prev, { baseName: base, subNames: cleanSubs, chunks }]);

        // 좌측에서 제거
        const remove = new Set(sorted);
        setAvailableIdx((prev) => prev.filter((i) => !remove.has(i)));
        setSelectedIdx([]);
        setSetBaseName('');
        setSubNames(['']);
    };

    // 우측 미리보기용 열 이름
    const previewRight = useMemo(() => {
        const items = [];

        // 그룹
        groups.forEach((g, gi) => {
            items.push({
                type: 'group',
                key: `group-${gi}`,
                title: g.name,
                cols: g.indices.map((i) => masterHeaders[i] ?? `열${i + 1}`)
            });
        });

        // 세트
        sets.forEach((s, si) => {
            const all = [];
            s.chunks.forEach((ch) => {
                ch.indices.forEach((colIdx, j) => {
                    const oldName = masterHeaders[colIdx] ?? `열${colIdx + 1}`;
                    const newName = ch.newHeaders[j];
                    all.push(`${oldName} → ${newName}`);
                });
            });
            items.push({
                type: 'set',
                key: `set-${si}`,
                title: s.baseName,
                cols: all
            });
        });

        return items;
    }, [groups, sets, masterHeaders]);

    // 내보내기: 선택된 컬럼만, 두 줄 헤더(상: 그룹/세트명, 하: 상세헤더) + 병합
    const handleDownload = () => {
        if (masterHeaders.length === 0) return;

        // 내보낼 열의 구성과 상/하단 헤더 만들기
        const topHeader = [];    // 상단(그룹/세트명)
        const bottomHeader = []; // 하단(상세 헤더명)
        const colPickers = [];   // 각 열에 대응하는 원본 인덱스와 새 이름

        // 그룹 먼저
        groups.forEach((g) => {
            g.indices.forEach((idx) => {
                topHeader.push(g.name);
                bottomHeader.push(masterHeaders[idx] ?? `열${idx + 1}`);
                colPickers.push({ srcIdx: idx });
            });
        });

        // 세트
        sets.forEach((s) => {
            s.chunks.forEach((ch) => {
                ch.indices.forEach((idx, j) => {
                    topHeader.push(s.baseName);
                    bottomHeader.push(ch.newHeaders[j]);
                    colPickers.push({ srcIdx: idx });
                });
            });
        });

        if (colPickers.length === 0) return;

        // 데이터 구성
        const out = [];
        out.push(topHeader);
        out.push(bottomHeader);
        dataRows.forEach((row) => {
            const newRow = colPickers.map(({ srcIdx }) => row[srcIdx]);
            out.push(newRow);
        });

        const ws = XLSX.utils.aoa_to_sheet(out);

        // 상단 헤더 병합: 같은 값이 연속되는 범위 병합
        ws['!merges'] = ws['!merges'] || [];
        let start = 0;
        for (let c = 1; c <= topHeader.length; c++) {
            if (c === topHeader.length || topHeader[c] !== topHeader[start]) {
                // [start, c-1] 범위 병합 (두 줄 헤더 중 1행만)
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
    };

    const leftHeaders = useMemo(
        () => availableIdx.map((i) => ({ i, name: masterHeaders[i] ?? `열${i + 1}` })),
        [availableIdx, masterHeaders]
    );

    const canMakeGroup = groupName.trim() && selectedIdx.length > 0;
    const canMakeSet =
        setBaseName.trim() &&
        subNames.map(s => s.trim()).filter(Boolean).length > 0 &&
        selectedIdx.length > 0 &&
        selectedIdx.length % subNames.map(s => s.trim()).filter(Boolean).length === 0;

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
                <Paper sx={{ width: '40%', p: 2, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    <Typography variant="h6" gutterBottom>1) 엑셀 업로드</Typography>
                    <Button variant="outlined" component="label" sx={{ mb: 2 }}>
                        엑셀 선택
                        <input
                            type="file"
                            hidden
                            accept=".xlsx"
                            onChange={(e) => e.target.files?.[0] && handleUpload(e.target.files[0])}
                        />
                    </Button>

                    <Divider sx={{ my: 2 }} />

                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
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

                    <Box sx={{ mt: 1, overflow: 'auto', flex: 1, border: '1px solid #eee', borderRadius: 1 }}>
                        <List dense>
                            {leftHeaders.map(({ i, name }) => (
                                <ListItem
                                    key={i}
                                    sx={{ py: 0.5, cursor: 'pointer' }}
                                    onClick={(e) => toggleSelect(i, e)}     // ← 리스트 아이템 클릭도 허용(체크박스 작은 것 대비 UX)
                                >
                                    <Checkbox
                                        size="small"
                                        checked={selectedIdx.includes(i)}
                                        onChange={(e) => toggleSelect(i, e)}  // ← 이벤트 전달
                                        onClick={(e) => e.stopPropagation()}  // ← 체크박스 클릭이 부모 onClick과 중복되지 않게
                                    />
                                    <Tooltip title={`열 인덱스: ${i + 1}`}>
                                        <Typography variant="body2">{name || `(빈 헤더) - 열${i + 1}`}</Typography>
                                    </Tooltip>
                                </ListItem>
                            ))}
                        </List>

                    </Box>

                    <Divider sx={{ my: 2 }} />

                    {/* 그룹 만들기 */}
                    <Typography variant="h6" gutterBottom>3) 그룹으로 묶기</Typography>
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

                    <Divider sx={{ my: 2 }} />

                    {/* 세트 만들기 */}
                    <Typography variant="h6" gutterBottom>4) 세트로 묶기</Typography>
                    <TextField
                        size="small"
                        placeholder="세트명 입력 (예: 경력사항)"
                        value={setBaseName}
                        onChange={(e) => setSetBaseName(e.target.value)}
                        fullWidth
                        sx={{ mb: 1 }}
                    />

                    <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
                        하위 항목명(세트의 1개 단위를 구성하는 컬럼명들)을 순서대로 입력하세요.
                    </Typography>

                    <Stack spacing={1} sx={{ maxHeight: 160, overflow: 'auto', mb: 1 }}>
                        {subNames.map((v, idx) => (
                            <Stack direction="row" spacing={1} key={idx}>
                                <TextField
                                    size="small"
                                    placeholder={`하위명 ${idx + 1}`}
                                    value={v}
                                    onChange={(e) => changeSubName(idx, e.target.value)}
                                    fullWidth
                                />
                                <Button
                                    variant="outlined"
                                    onClick={() => removeSubName(idx)}
                                    disabled={subNames.length === 1}
                                >
                                    삭제
                                </Button>
                            </Stack>
                        ))}
                    </Stack>
                    <Stack direction="row" spacing={1} sx={{ mb: 1 }}>
                        <Button variant="outlined" onClick={addSubName}>하위명 추가</Button>
                        <Chip
                            label={`선택된 열: ${selectedIdx.length}개 / 세트 단위: ${subNames.map(s => s.trim()).filter(Boolean).length || 0}`}
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
