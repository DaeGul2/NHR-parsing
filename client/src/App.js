import React, { useState, useEffect } from 'react';
import {
  Container,
  Typography,
  Divider,
  MenuItem,
  Select,
  InputLabel,
  FormControl,
  Button,
  Box,
  FormControlLabel,
  Checkbox,
  Paper,
  Grid,
  Chip,
  Stack,
  Tooltip
} from '@mui/material';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

import FileUploader from './components/FileUploader';
import SheetConfigurator from './components/SheetConfigurator';
import GenerateButton from './components/GenerateButton';
import { generateStep2Excel } from './utils/generateStep2Excel';
import NhrTransformModal from './components/NhrTransformModal';

// ✅ 컬럼 병합 페이지 임포트
import ColumnMergePage from './ColumnMergePage';

// 간단한 드래그앤드롭 리스트 컴포넌트 (스타일 개선)
function DraggableList({ items, onOrderChange }) {
  const [dragIndex, setDragIndex] = useState(null);

  const handleDragStart = (index) => {
    setDragIndex(index);
  };

  const handleDragOver = (index, e) => {
    e.preventDefault();
  };

  const handleDrop = (index) => {
    if (dragIndex === null) return;
    const newItems = [...items];
    const [removed] = newItems.splice(dragIndex, 1);
    newItems.splice(index, 0, removed);
    onOrderChange(newItems);
    setDragIndex(null);
  };

  return (
    <Box
      sx={{
        borderRadius: 2,
        border: '1px dashed',
        borderColor: 'divider',
        p: 1.5,
        background:
          'linear-gradient(135deg, rgba(144,202,249,0.06), rgba(206,147,216,0.06))'
      }}
    >
      <Stack direction="row" flexWrap="wrap" gap={1}>
        {items.map((item, index) => (
          <Box
            key={index}
            draggable
            onDragStart={() => handleDragStart(index)}
            onDragOver={(e) => handleDragOver(index, e)}
            onDrop={() => handleDrop(index)}
            sx={{
              px: 1.5,
              py: 0.75,
              borderRadius: 999,
              border: '1px solid',
              borderColor: dragIndex === index ? 'primary.main' : 'divider',
              backgroundColor:
                dragIndex === index ? 'primary.light' : 'background.paper',
              boxShadow: dragIndex === index ? 2 : 0,
              cursor: 'grab',
              fontSize: 13,
              display: 'inline-flex',
              alignItems: 'center',
              gap: 0.75
            }}
          >
            <Box
              component="span"
              sx={{
                width: 8,
                height: 8,
                borderRadius: '50%',
                bgcolor: 'primary.main'
              }}
            />
            {item}
          </Box>
        ))}
      </Stack>
    </Box>
  );
}

function App() {
  // ✅ 컬럼 병합 페이지 토글
  const [showColumnMerge, setShowColumnMerge] = useState(false);

  // Step1 관련 상태
  const [headerRow, setHeaderRow] = useState([]);
  const [rows, setRows] = useState([]);
  const [groups, setGroups] = useState([]);
  const [selectedGroups, setSelectedGroups] = useState([]);
  // 기준 컬럼은 파일마다 다를 수 있으므로 유동적으로 선택 (예: 기본정보)
  const [idCols, setIdCols] = useState({ 지원자번호: '', 지원직무: '', 이름: '' });
  const [step1Workbook, setStep1Workbook] = useState(null);

  // Step2 관련 상태
  const [step2Sheets, setStep2Sheets] = useState([]); // step1 결과물 중 세로화할 시트 선택
  const [groupSets, setGroupSets] = useState({});       // { sheetName: [ [set1], [set2], ... ] }
  const [activeGroupSet, setActiveGroupSet] = useState({}); // { sheetName: 선택된 세트의 index }
  const [sortRules, setSortRules] = useState({});         // { sheetName: { key: '정렬컬럼', method: 'desc'|'asc'|'alpha' } }
  const [groupColumnOrder, setGroupColumnOrder] = useState({});

  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);

  // NHR 변환 모달
  const [openNHR, setOpenNHR] = useState(false);

  useEffect(() => {
    if (step1Workbook) {
      const sheets = step1Workbook.SheetNames.filter(name => name !== 'rawdata');
      sheets.forEach(sheet => autoDetectGroupSets(sheet));
    }
  }, [step1Workbook]);

  // 사용자가 선택한 기준 컬럼 역할은 동일하지만, step1 결과물에선 "컬럼명"만 사용됨
  const baseColumns = ['지원자번호', '지원직무', '이름'];

  // 파일 업로드 (Step1 원본 엑셀)
  const handleUpload = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const merges = sheet['!merges'] || [];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      merges.forEach(({ s, e }) => {
        const row = s.r;
        const val = raw[row][s.c];
        for (let i = s.c; i <= e.c; i++) raw[row][i] = val;
      });
      setHeaderRow(raw[0]);
      setRows(raw.slice(1));
      const uniqueGroups = [...new Set(raw[0].filter(Boolean))];
      setGroups(uniqueGroups);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleIdColChange = (key) => (e) => {
    setIdCols({ ...idCols, [key]: parseInt(e.target.value, 10) });
  };

  // Step1: 엑셀 생성 및 다운로드, 내부에 Step2용 워크북 저장
  const generateStep1 = () => {
    setLoading(true); // ✅ 로딩 시작
    const wb = XLSX.utils.book_new();
    const groupMap = {};
    headerRow.forEach((group, idx) => {
      if (!group) return;
      if (!groupMap[group]) groupMap[group] = [];
      groupMap[group].push(idx);
    });
    selectedGroups.forEach((group) => {
      const indices = groupMap[group];
      if (!indices || indices.length === 0) return;
      const groupRows = rows.map((row) => {
        // 기준 컬럼은 그대로 유지 (빈 값은 빈 문자열)
        const base = [
          row[idCols['지원자번호']] !== undefined ? row[idCols['지원자번호']] : '',
          row[idCols['지원직무']] !== undefined ? row[idCols['지원직무']] : '',
          row[idCols['이름']] !== undefined ? row[idCols['이름']] : ''
        ];
        const groupData = indices.map(i => row[i]);
        return [...base, ...groupData];
      });
      const ws = XLSX.utils.aoa_to_sheet(groupRows);
      XLSX.utils.book_append_sheet(wb, ws, group);
    });
    const wsRaw = XLSX.utils.aoa_to_sheet([headerRow, ...rows]);
    XLSX.utils.book_append_sheet(wb, wsRaw, 'rawdata');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    saveAs(blob, 'step1.xlsx');
    const newWb = XLSX.read(wbout, { type: 'array' });
    setStep1Workbook(newWb);
    // Step2 대상 시트: rawdata 제외
    setStep2Sheets(newWb.SheetNames.filter(name => name !== 'rawdata'));

    setLoading(false); // ✅ 로딩 종료
  };

  // Step1에서 선택한 기준 컬럼의 인덱스를 통해, step1 결과물에 나온 "컬럼명"(서브컬럼)만 추출
  const getIdKeyNames = () => {
    return baseColumns.map(key => {
      const idx = idCols[key];
      return rows[0]?.[idx] || '이름없음';
    });
  };

  // Step2: 시트 선택 토글
  const handleStep2SheetToggle = (sheet) => {
    setStep2Sheets(prev =>
      prev.includes(sheet) ? prev.filter(s => s !== sheet) : [...prev, sheet]
    );
  };

  // Step2: 각 시트별 자동 그룹 감지
  // 조건: 반복되는 컬럼명이 문자+숫자로 연속되는 경우 그룹화 (기본정보, 연번 제외)
  const autoDetectGroupSets = (sheetName) => {
    const sheet = step1Workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const allCols = Object.keys(data[0] || {});
    const repeatCols = allCols.filter(c => !baseColumns.includes(c) && c !== '연번');
    const sets = {};
    repeatCols.forEach(col => {
      const match = col.match(/(\D+)(\d+)$/);
      if (match) {
        const key = match[2];
        if (!sets[key]) sets[key] = [];
        sets[key].push(col);
      }
    });
    const grouped = Object.values(sets);
    setGroupSets(prev => ({ ...prev, [sheetName]: grouped }));
    if (grouped.length > 0) {
      setActiveGroupSet(prev => ({ ...prev, [sheetName]: 0 }));
      // 자동 감지된 첫 세트에서 숫자 제거한 값들
      const detected = grouped[0].map(col => col.replace(/\d+$/, ''));
      setGroupColumnOrder(prev => ({ ...prev, [sheetName]: detected }));
    }
  };

  // Step2: 그룹 컬럼 순서 재정렬 (드래그앤드롭)
  const handleGroupColumnOrderChange = (sheet, newOrder) => {
    setGroupColumnOrder(prev => ({ ...prev, [sheet]: newOrder }));
  };

  // 최종 컬럼 순서는: 기준 컬럼(실제 step1 결과물의 컬럼명), '연번', 그리고 재정렬된 그룹 컬럼 순서
  const getFinalColumnOrder = (sheet) => {
    return [...getIdKeyNames(), '연번', ...(groupColumnOrder[sheet] || [])];
  };

  // Step2: 정렬 기준 및 방식 선택 (정렬 안 함 선택 가능)
  const handleSortRuleChange = (sheet, key, method) => {
    setSortRules(prev => ({ ...prev, [sheet]: { key, method } }));
  };

  // Step2: 최종 Step2 엑셀 생성 (generateStep2Excel 호출)
  const handleGenerateStep2 = () => {
    setGenerating(true); // ✅ 로딩 시작
    // 각 시트에 대해 활성 그룹 세트만 전달하고, 최종 컬럼 순서를 계산하여 전달
    const activeGroupSetsForExcel = {};
    const finalColumnOrders = {};
    step2Sheets.forEach(sheet => {
      if (groupSets[sheet] && activeGroupSet[sheet] !== undefined) {
        activeGroupSetsForExcel[sheet] = [groupSets[sheet][activeGroupSet[sheet]]];
      }
      finalColumnOrders[sheet] = getFinalColumnOrder(sheet);
    });
    generateStep2Excel({
      workbook: step1Workbook,
      selectedSheets: step2Sheets,
      groupSets: activeGroupSetsForExcel,
      columnOrders: finalColumnOrders,
      sortRules,
      // step1에서 선택한 기준컬럼(실제 서브컬럼명)을 그대로 사용
      idKeys: getIdKeyNames()
    });
    setGenerating(false); // ✅ 로딩 종료
  };

  // ✅ '컬럼 병합하기' 버튼을 누르면 병합 전용 화면으로 전환
  if (showColumnMerge) {
    return (
      <Container maxWidth="lg" sx={{ py: 4 }}>
        <Paper
          elevation={3}
          sx={{
            p: 2,
            mb: 3,
            borderRadius: 3,
            display: 'flex',
            alignItems: 'center',
            gap: 1.5
          }}
        >
          <Button variant="outlined" onClick={() => setShowColumnMerge(false)}>
            뒤로 가기
          </Button>
          <Typography variant="h6" sx={{ fontWeight: 700 }}>
            컬럼 병합 페이지
          </Typography>
        </Paper>
        <Paper
          elevation={1}
          sx={{
            p: 3,
            borderRadius: 3,
            backgroundColor: 'background.paper'
          }}
        >
          <ColumnMergePage />
        </Paper>
      </Container>
    );
  }

  return (
    <Container
      maxWidth="lg"
      sx={{
        py: 4,
        pb: 6
      }}
    >
      {/* 헤더 카드 */}
      <Paper
        elevation={3}
        sx={{
          p: 3,
          mb: 3,
          borderRadius: 3,
          background:
            'linear-gradient(135deg, rgba(33,150,243,0.12), rgba(156,39,176,0.12))',
          border: '1px solid',
          borderColor: 'divider'
        }}
      >
        <Box
          sx={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'flex-start',
            gap: 2,
            flexWrap: 'wrap'
          }}
        >
          <Box>
            <Typography variant="h4" sx={{ fontWeight: 800, mb: 1 }}>
              시트 분리 & 컬럼 세로화 도구
            </Typography>
            <Typography variant="body2" color="text.secondary">
              지원자 기본 정보는 유지하면서, 그룹별 컬럼을 자동으로 나누고 세로화할 수 있는 작업용
              화면입니다.
            </Typography>
          </Box>
          <Stack direction="row" spacing={1} flexWrap="wrap" justifyContent="flex-end">
            <Chip label="Step1 · 시트 분리" size="small" color="primary" variant="outlined" />
            <Chip label="Step2 · 세로화 / 정렬" size="small" color="secondary" variant="outlined" />
            <Chip label="NHR 포맷 변환" size="small" variant="outlined" />
          </Stack>
        </Box>
      </Paper>

      {/* 상단 액션 영역 */}
      <Paper
        elevation={1}
        sx={{
          p: 2,
          mb: 3,
          borderRadius: 3,
          display: 'flex',
          flexWrap: 'wrap',
          alignItems: 'center',
          gap: 1.5,
          backgroundColor: 'background.paper'
        }}
      >
        <Tooltip title="NHR 양식으로 가공된 엑셀 파일이 필요할 때 사용">
          <Button variant="outlined" onClick={() => setOpenNHR(true)}>
            NHR 형식으로 엑셀 바꾸기
          </Button>
        </Tooltip>

        <Tooltip title="원본 엑셀 업로드 (Step1 기준)">
          <Box>
            <FileUploader onUpload={handleUpload} />
          </Box>
        </Tooltip>

        {/* ✅ 컬럼 병합하기 버튼 */}
        <Tooltip title="별도 페이지에서 복수 컬럼을 한 컬럼으로 병합">
          <Button variant="contained" color="primary" onClick={() => setShowColumnMerge(true)}>
            컬럼 병합하기
          </Button>
        </Tooltip>

        <Box sx={{ flexGrow: 1 }} />

        {(loading || generating) && (
          <Typography variant="body2" color="text.secondary">
            {loading ? 'Step1 엑셀 생성 중...' : 'Step2 엑셀 생성 중...'}
          </Typography>
        )}
      </Paper>

      {/* Step1 영역 */}
      {groups.length > 0 && (
        <Paper
          elevation={1}
          sx={{
            p: 3,
            mb: 3,
            borderRadius: 3,
            backgroundColor: 'background.paper'
          }}
        >
          <Box sx={{ mb: 2, display: 'flex', justifyContent: 'space-between', gap: 2 }}>
            <Box>
              <Typography variant="h6" sx={{ fontWeight: 700 }}>
                Step1 · 기준 컬럼 선택 & 시트 분리
              </Typography>
              <Typography variant="body2" color="text.secondary">
                지원자번호 / 지원직무 / 이름에 해당하는 열을 지정한 후, 그룹 단위로 시트를 나눕니다.
              </Typography>
            </Box>
            <Chip
              label="원본 시트 → 그룹별 시트"
              size="small"
              color="primary"
              variant="outlined"
            />
          </Box>

          <Grid container spacing={2}>
            <Grid item xs={12} md={6}>
              <Typography variant="subtitle2" sx={{ mb: 1 }}>
                기준 컬럼 선택
              </Typography>
              {['지원자번호', '지원직무', '이름'].map((key) => (
                <FormControl fullWidth sx={{ my: 1 }} key={key} size="small">
                  <InputLabel>{key}</InputLabel>
                  <Select value={idCols[key]} label={key} onChange={handleIdColChange(key)}>
                    {headerRow.map((group, idx) => {
                      const secondRowValue = rows[0]?.[idx] || '(값 없음)';
                      return (
                        <MenuItem key={idx} value={idx}>
                          {`${group || '그룹없음'}: ${secondRowValue}`}
                        </MenuItem>
                      );
                    })}
                  </Select>
                </FormControl>
              ))}
            </Grid>

            <Grid item xs={12} md={6}>
              <Typography variant="subtitle2" sx={{ mb: 1 }}>
                시트로 분리할 그룹 선택
              </Typography>
              <SheetConfigurator
                headers={headerRow}
                selected={selectedGroups}
                setSelected={setSelectedGroups}
              />
            </Grid>
          </Grid>

          <Divider sx={{ my: 2 }} />

          <Box sx={{ display: 'flex', justifyContent: 'flex-end' }}>
            <GenerateButton onClick={generateStep1} />
          </Box>
        </Paper>
      )}

      {/* Step2 영역 */}
      {step1Workbook && (
        <Paper
          elevation={1}
          sx={{
            p: 3,
            borderRadius: 3,
            backgroundColor: 'background.paper'
          }}
        >
          <Box sx={{ mb: 2, display: 'flex', justifyContent: 'space-between', gap: 2 }}>
            <Box>
              <Typography variant="h6" sx={{ fontWeight: 700 }}>
                Step2 · 시트 세로화 설정
              </Typography>
              <Typography variant="body2" color="text.secondary">
                사용할 시트를 선택하고, 반복되는 컬럼을 세트로 묶어 세로화한 뒤 정렬 기준까지 지정할 수
                있습니다.
              </Typography>
            </Box>
            <Chip
              label="그룹 컬럼 → 세로 레코드"
              size="small"
              color="secondary"
              variant="outlined"
            />
          </Box>

          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            {step1Workbook.SheetNames.filter(name => name !== 'rawdata').map((sheetName) => (
              <Paper
                key={sheetName}
                variant="outlined"
                sx={{
                  p: 2,
                  borderRadius: 2,
                  backgroundColor: 'background.default'
                }}
              >
                <Box
                  sx={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 1
                  }}
                >
                  <FormControlLabel
                    control={
                      <Checkbox
                        checked={step2Sheets.includes(sheetName)}
                        onChange={() => handleStep2SheetToggle(sheetName)}
                      />
                    }
                    label={
                      <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
                        {sheetName}
                      </Typography>
                    }
                  />
                  <Stack direction="row" spacing={1} alignItems="center">
                    <Typography variant="caption" color="text.secondary">
                      기준 컬럼 순서:
                    </Typography>
                    <Stack direction="row" spacing={0.5}>
                      {getFinalColumnOrder(sheetName).slice(0, 3).map(col => (
                        <Chip key={col} label={col} size="small" variant="outlined" />
                      ))}
                      {getFinalColumnOrder(sheetName).length > 3 && (
                        <Chip
                          label={`+${getFinalColumnOrder(sheetName).length - 3}`}
                          size="small"
                          variant="outlined"
                        />
                      )}
                    </Stack>
                  </Stack>
                </Box>

                {step2Sheets.includes(sheetName) && (
                  <Box sx={{ mt: 2 }}>
                    <Box sx={{ display: 'flex', justifyContent: 'space-between', mb: 1 }}>
                      <Typography variant="subtitle2">그룹 세트 설정</Typography>
                      <Button size="small" onClick={() => autoDetectGroupSets(sheetName)}>
                        자동 그룹 감지
                      </Button>
                    </Box>

                    {groupSets[sheetName] && groupSets[sheetName].length > 0 && (
                      <Grid container spacing={2}>
                        <Grid item xs={12} md={6}>
                          <FormControl fullWidth size="small">
                            <InputLabel>활성 그룹 세트 선택</InputLabel>
                            <Select
                              value={
                                activeGroupSet[sheetName] !== undefined
                                  ? activeGroupSet[sheetName]
                                  : ''
                              }
                              label="활성 그룹 세트 선택"
                              onChange={(e) => {
                                const index = parseInt(e.target.value, 10);
                                setActiveGroupSet(prev => ({ ...prev, [sheetName]: index }));
                                const newOrder = groupSets[sheetName][index].map(col =>
                                  col.replace(/\d+$/, '')
                                );
                                setGroupColumnOrder(prev => ({ ...prev, [sheetName]: newOrder }));
                              }}
                            >
                              {groupSets[sheetName].map((set, idx) => (
                                <MenuItem key={idx} value={idx}>
                                  세트 {idx + 1}: [ {set.join(', ')} ]
                                </MenuItem>
                              ))}
                            </Select>
                          </FormControl>

                          {groupColumnOrder[sheetName] && (
                            <Box sx={{ mt: 2 }}>
                              <Typography variant="body2" sx={{ mb: 0.5 }}>
                                최종 그룹 컬럼 순서 (기본정보 뒤에 배치됨)
                              </Typography>
                              <DraggableList
                                items={groupColumnOrder[sheetName]}
                                onOrderChange={(newOrder) =>
                                  handleGroupColumnOrderChange(sheetName, newOrder)
                                }
                              />
                            </Box>
                          )}
                        </Grid>

                        <Grid item xs={12} md={6}>
                          <Typography variant="subtitle2" sx={{ mb: 1 }}>
                            정렬 옵션
                          </Typography>
                          <FormControl fullWidth size="small" sx={{ mb: 1 }}>
                            <InputLabel>정렬 기준 컬럼</InputLabel>
                            <Select
                              value={sortRules[sheetName]?.key || ''}
                              label="정렬 기준 컬럼"
                              onChange={(e) =>
                                handleSortRuleChange(
                                  sheetName,
                                  e.target.value,
                                  sortRules[sheetName]?.method || ''
                                )
                              }
                            >
                              <MenuItem value="">정렬 안 함</MenuItem>
                              {getFinalColumnOrder(sheetName).map(col => (
                                <MenuItem key={col} value={col}>
                                  {col}
                                </MenuItem>
                              ))}
                            </Select>
                          </FormControl>
                          <FormControl fullWidth size="small">
                            <InputLabel>정렬 방식</InputLabel>
                            <Select
                              value={sortRules[sheetName]?.method || ''}
                              label="정렬 방식"
                              onChange={(e) =>
                                handleSortRuleChange(
                                  sheetName,
                                  sortRules[sheetName]?.key || '',
                                  e.target.value
                                )
                              }
                            >
                              <MenuItem value="">정렬 안 함</MenuItem>
                              <MenuItem value="desc">내림차순</MenuItem>
                              <MenuItem value="asc">오름차순</MenuItem>
                              <MenuItem value="alpha">가나다순</MenuItem>
                            </Select>
                          </FormControl>
                        </Grid>
                      </Grid>
                    )}
                  </Box>
                )}
              </Paper>
            ))}

            <Box sx={{ display: 'flex', justifyContent: 'flex-end', mt: 1 }}>
              <Button
                variant="contained"
                color="secondary"
                onClick={handleGenerateStep2}
                disabled={generating}
              >
                Step2 엑셀 생성
              </Button>
            </Box>
          </Box>
        </Paper>
      )}

      {/* NHR 변환 모달 */}
      <NhrTransformModal open={openNHR} onClose={() => setOpenNHR(false)} />
    </Container>
  );
}

export default App;
