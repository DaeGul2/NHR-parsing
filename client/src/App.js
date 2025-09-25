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
  Checkbox
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
    <Box sx={{ border: '1px dashed #aaa', p: 2, borderRadius: 1 }}>
      {items.map((item, index) => (
        <Box
          key={index}
          draggable
          onDragStart={() => handleDragStart(index)}
          onDragOver={(e) => handleDragOver(index, e)}
          onDrop={() => handleDrop(index)}
          sx={{
            p: 1,
            border: '1px solid #ccc',
            mb: 1,
            borderRadius: 1,
            backgroundColor: '#f7f7f7',
            cursor: 'grab'
          }}
        >
          {item}
        </Box>
      ))}
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
      <Container sx={{ py: 2 }}>
        <Box sx={{ display: 'flex', gap: 1, mb: 2 }}>
          <Button variant="outlined" onClick={() => setShowColumnMerge(false)}>
            뒤로 가기
          </Button>
        </Box>
        <Typography variant="h5" gutterBottom>
          컬럼 병합 페이지
        </Typography>
        <ColumnMergePage />
      </Container>
    );
  }

  return (
    <Container>
      <Typography variant="h4" gutterBottom>
        시트 분리 및 컬럼 세로화 작업
      </Typography>

      {/* 상단 버튼 줄: NHR 변환 + 엑셀 업로드 + ✅ 컬럼 병합하기 */}
      <Box sx={{ display: 'flex', gap: 1, mb: 2 }}>
        <Button variant="outlined" onClick={() => setOpenNHR(true)}>
          nhr 형식으로 엑셀 바꾸기
        </Button>
        <FileUploader onUpload={handleUpload} />
        {/* ✅ 컬럼 병합하기 버튼 */}
        <Button variant="contained" color="primary" onClick={() => setShowColumnMerge(true)}>
          컬럼 병합하기
        </Button>
      </Box>

      {/* Step1 영역 */}
      {groups.length > 0 && (
        <>
          <Divider sx={{ my: 2 }} />
          <Typography variant="h6">📌 기준 컬럼 선택 (Step1)</Typography>
          {['지원자번호', '지원직무', '이름'].map((key) => (
            <FormControl fullWidth sx={{ my: 1 }} key={key}>
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
          <Divider sx={{ my: 2 }} />
          <SheetConfigurator
            headers={headerRow}
            selected={selectedGroups}
            setSelected={setSelectedGroups}
          />
          <GenerateButton onClick={generateStep1} />
        </>
      )}

      {/* Step2 영역 */}
      {step1Workbook && (
        <>
          <Divider sx={{ my: 3 }} />
          <Typography variant="h5">Step2: 시트 세로화 설정</Typography>
          {step1Workbook.SheetNames.filter(name => name !== 'rawdata').map((sheetName) => (
            <Box key={sheetName} sx={{ my: 2, border: '1px solid #ccc', p: 2 }}>
              <FormControlLabel
                control={
                  <Checkbox
                    checked={step2Sheets.includes(sheetName)}
                    onChange={() => handleStep2SheetToggle(sheetName)}
                  />
                }
                label={sheetName}
              />
              {step2Sheets.includes(sheetName) && (
                <>
                  <Box sx={{ mt: 1 }}>
                    <Button size="small" onClick={() => autoDetectGroupSets(sheetName)}>
                      자동 그룹 감지
                    </Button>
                  </Box>
                  {groupSets[sheetName] && groupSets[sheetName].length > 0 && (
                    <Box sx={{ mt: 1 }}>
                      <FormControl fullWidth>
                        <InputLabel>활성 그룹 세트 선택</InputLabel>
                        <Select
                          value={activeGroupSet[sheetName] !== undefined ? activeGroupSet[sheetName] : ''}
                          label="활성 그룹 세트 선택"
                          onChange={(e) => {
                            const index = parseInt(e.target.value, 10);
                            setActiveGroupSet(prev => ({ ...prev, [sheetName]: index }));
                            const newOrder = groupSets[sheetName][index].map(col => col.replace(/\d+$/, ''));
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
                        <>
                          <Typography variant="body2" sx={{ mt: 1 }}>
                            최종 그룹 컬럼 순서 (기본정보 뒤에 배치됨):
                          </Typography>
                          <DraggableList
                            items={groupColumnOrder[sheetName]}
                            onOrderChange={(newOrder) => handleGroupColumnOrderChange(sheetName, newOrder)}
                          />
                        </>
                      )}
                    </Box>
                  )}
                  <Box sx={{ mt: 2 }}>
                    <Typography variant="body2">정렬 기준 선택 (선택 사항)</Typography>
                    <FormControl fullWidth sx={{ mt: 1 }}>
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
                          <MenuItem key={col} value={col}>{col}</MenuItem>
                        ))}
                      </Select>
                    </FormControl>
                    <FormControl fullWidth sx={{ mt: 1 }}>
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
                  </Box>
                </>
              )}
            </Box>
          ))}
          <Button variant="contained" color="secondary" onClick={handleGenerateStep2}>
            Step2 엑셀 생성
          </Button>
        </>
      )}

      {/* NHR 변환 모달 */}
      <NhrTransformModal open={openNHR} onClose={() => setOpenNHR(false)} />
    </Container>
  );
}

export default App;
