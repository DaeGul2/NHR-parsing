import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

/** ====== 세트 컬럼 패턴: "<세트명><번호> - <서브필드>" ====== */
const SET_COLUMN_REGEX = /^(.*?)(\d+)\s*[-–—:]\s*(.+)$/u;
const norm = (s) => (typeof s === "string" ? s.replace(/\s+/g, " ").trim() : s);

/** XLSX/CSV 첫 시트 표 전체 읽기: headers + rows(데이터 행들) */
async function readTableFromFile(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
  const headers = (aoa[0] || []).map(norm).filter((h) => !!h);
  const rows = (aoa.slice(1) || []).map((r) => r.map((v) => (v == null ? "" : v)));
  return { headers, rows };
}

/** 헤더에서 단일/세트 분류 */
function classifyHeaders(headers) {
  const singles = [];
  const sets = {};
  for (const raw of headers) {
    const col = norm(raw);
    if (!col) continue;
    const m = col.match(SET_COLUMN_REGEX);
    if (m) {
      const setName = norm(m[1]);
      const idx = norm(m[2]);
      const field = norm(m[3]);
      if (!sets[setName]) sets[setName] = { indices: [], fields: [], byIndex: {} };
      if (!sets[setName].byIndex[idx]) {
        sets[setName].byIndex[idx] = [];
        sets[setName].indices.push(idx);
      }
      if (!sets[setName].fields.includes(field)) sets[setName].fields.push(field);
      sets[setName].byIndex[idx].push(field);
    } else {
      singles.push(col);
    }
  }
  Object.values(sets).forEach((g) => g.indices.sort((a, b) => Number(a) - Number(b)));
  return { singles, sets };
}

/** 병합 미리보기용 2행 헤더 생성 (표시는 ‘세트명 - 필드명{N}’) */
function buildMergedTwoRowHeaders(stateSingles, stateSets) {
  const row2 = [];
  const groups = [];

  // 단일
  for (const s of stateSingles) row2.push(s);
  if (stateSingles.length > 0) groups.push({ title: "단일컬럼", span: stateSingles.length });

  // 세트 (표시 라벨만 새 포맷)
  Object.entries(stateSets).forEach(([setName, g]) => {
    const fields = g.fields;
    const indices = g.indices;
    const span = fields.length * indices.length;
    if (span === 0) return;
    for (const idx of indices) {
      for (const f of fields) {
        // 미리보기/다운로드 헤더 라벨
        row2.push(`${setName} - ${f}${idx}`);
      }
    }
    groups.push({ title: setName, span });
  });

  const row1 = groups.map((g) => ({ title: g.title, colSpan: g.span }));
  return { row1, row2 };
}

/** === 데이터까지 포함한 최종 AOA 생성 ===
 *  표시 헤더는 (세트명 - 필드명{N})로 쓰되,
 *  데이터 조회는 원본 키 (세트명{N} - 필드명)로 해야 한다!
 */
function buildExportAOA(origHeaders, origRows, singles, sets) {
  const { row1, row2 } = buildMergedTwoRowHeaders(singles, sets);

  // 원본 헤더 인덱스 맵(원본 그대로)
  const headerIndexMap = new Map();
  origHeaders.forEach((h, i) => headerIndexMap.set(h, i));

  // 1행(병합용 그룹 타이틀) 확장
  const row1Expanded = [];
  row1.forEach((g) => {
    for (let i = 0; i < g.colSpan; i++) row1Expanded.push(g.title);
  });

  // 데이터 행 매핑: 표시 라벨은 새 포맷이지만, 값 찾기는 '원본 포맷'으로!
  const body = origRows.map((r) => {
    const out = [];
    // 단일
    for (const s of singles) {
      const idx = headerIndexMap.get(s);
      let val = idx != null ? r[idx] : "";
      if (val === "-") val = "";
      out.push(val);
    }
    // 세트
    Object.entries(sets).forEach(([setName, g]) => {
      for (const idxStr of g.indices) {
        for (const f of g.fields) {
          // ✅ 값 찾기용 '원본 헤더 키'
          const originalKey = `${setName}${idxStr} - ${f}`;
          const colIdx = headerIndexMap.get(originalKey);
          let val = colIdx != null ? r[colIdx] : "";
          if (val === "-") val = "";
          out.push(val);
        }
      }
    });
    return out;
  });

  const aoa = [row1Expanded, row2, ...body];

  // 1행 병합정보
  const merges = [];
  let c = 0;
  row1.forEach((g) => {
    const start = c;
    const end = c + g.colSpan - 1;
    if (end > start) merges.push({ s: { r: 0, c: start }, e: { r: 0, c: end } });
    c += g.colSpan;
  });

  return { aoa, merges, fileName: "merged.xlsx" };
}

export default function ColumnMergePage() {
  /** 원본표 */
  const [origHeaders, setOrigHeaders] = useState([]);
  const [origRows, setOrigRows] = useState([]);

  /** 원본 분류 & 작업 상태 */
  const [orig, setOrig] = useState({ singles: [], sets: {} });
  const [singles, setSingles] = useState([]);
  const [sets, setSets] = useState({});

  /** 좌/우 UI 선택 상태 */
  const [singleSelection, setSingleSelection] = useState(new Set());
  const [openSetNames, setOpenSetNames] = useState(new Set());
  const [setSelection, setSetSelection] = useState(new Set());
  const [fieldSelection, setFieldSelection] = useState({});

  const fileInputRef = useRef(null);

  /** 파일 업로드 */
  const onFile = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;

    const { headers, rows } = await readTableFromFile(f);
    const cls = classifyHeaders(headers);

    setOrigHeaders(headers);
    setOrigRows(rows);

    setOrig(cls);
    setSingles(cls.singles);
    setSets(JSON.parse(JSON.stringify(cls.sets || {})));

    setSingleSelection(new Set());
    setOpenSetNames(new Set());
    setSetSelection(new Set());
    setFieldSelection({});
  };

  /** 좌측 토글/삭제/리셋 */
  const toggleSingle = (name) => {
    setSingleSelection((prev) => {
      const n = new Set(prev);
      n.has(name) ? n.delete(name) : n.add(name);
      return n;
    });
  };
  const deleteSingles_SelectedOnly = () => {
    if (singleSelection.size === 0) return;
    setSingles((prev) => prev.filter((s) => !singleSelection.has(s)));
    setSingleSelection(new Set());
  };
  const deleteSingles_ExceptSelected = () => {
    if (singleSelection.size === 0) {
      setSingles([]);
    } else {
      setSingles((prev) => prev.filter((s) => singleSelection.has(s)));
    }
    setSingleSelection(new Set());
  };
  const resetSingles = () => {
    setSingles(orig.singles || []);
    setSingleSelection(new Set());
  };

  /** 우측 아코디언/세트 선택/필드 선택/삭제/리셋 */
  const toggleSetOpen = (setName) => {
    setOpenSetNames((prev) => {
      const n = new Set(prev);
      n.has(setName) ? n.delete(setName) : n.add(setName);
      return n;
    });
  };
  const toggleSetSelected = (setName) => {
    setSetSelection((prev) => {
      const n = new Set(prev);
      n.has(setName) ? n.delete(setName) : n.add(setName);
      return n;
    });
  };
  const toggleFieldSelected = (setName, field) => {
    setFieldSelection((prev) => {
      const next = { ...prev };
      const curr = new Set(prev[setName] || []);
      curr.has(field) ? curr.delete(field) : curr.add(field);
      next[setName] = curr;
      return next;
    });
  };
  const deleteSets_SelectedOnly = () => {
    if (setSelection.size === 0) return;
    setSets((prev) => {
      const copy = { ...prev };
      for (const name of setSelection) delete copy[name];
      return copy;
    });
    setSetSelection(new Set());
  };
  const deleteSets_ExceptSelected = () => {
    setSets((prev) => {
      if (setSelection.size === 0) return {};
      const keep = {};
      Object.entries(prev).forEach(([name, v]) => {
        if (setSelection.has(name)) keep[name] = v;
      });
      return keep;
    });
    setSetSelection(new Set());
  };
  const deleteSelectedFieldsInSets = () => {
    setSets((prev) => {
      const copy = JSON.parse(JSON.stringify(prev));
      Object.entries(fieldSelection).forEach(([setName, selSet]) => {
        if (!copy[setName]) return;
        const selectedFields = Array.from(selSet || []);
        if (selectedFields.length === 0) return;

        copy[setName].fields = copy[setName].fields.filter((f) => !selectedFields.includes(f));
        Object.entries(copy[setName].byIndex).forEach(([idx, arr]) => {
          copy[setName].byIndex[idx] = arr.filter((f) => !selectedFields.includes(f));
        });
      });
      return copy;
    });
    setFieldSelection({});
  };
  const resetSets = () => {
    setSets(JSON.parse(JSON.stringify(orig.sets || {})));
    setSetSelection(new Set());
    setFieldSelection({});
  };

  /** 병합 미리보기 (2행 헤더) */
  const merged = useMemo(() => buildMergedTwoRowHeaders(singles, sets), [singles, sets]);

  /** 병합 + 데이터 포함하여 엑셀 다운로드 */
  const handleExportMerged = () => {
    if (!origHeaders.length) {
      alert("먼저 엑셀을 업로드해주세요.");
      return;
    }
    const hasAny =
      singles.length > 0 ||
      Object.values(sets).some((g) => g.fields.length * g.indices.length > 0);

    if (!hasAny) {
      alert("병합할 컬럼이 없습니다.");
      return;
    }

    const { aoa, merges, fileName } = buildExportAOA(origHeaders, origRows, singles, sets);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"] = merges;
    XLSX.utils.book_append_sheet(wb, ws, "merged");

    const wbout = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    saveAs(new Blob([wbout], { type: "application/octet-stream" }), fileName);
  };

  /** 병합 미리보기 테이블 */
  const MergedPreview = () => {
    const { row1, row2 } = merged;
    const totalCols = row2.length;

    if (totalCols === 0) {
      return (
        <div className="preview empty">
          <div>미리보기: 남은 컬럼이 없습니다.</div>
        </div>
      );
    }

    return (
      <div className="preview">
        <table className="merged-table">
          <thead>
            <tr>
              {row1.map((g, i) => (
                <th key={i} colSpan={g.colSpan} className="lvl1">
                  {g.title}
                </th>
              ))}
            </tr>
            <tr>
              {row2.map((h, i) => (
                <th key={i} className="lvl2">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
        </table>
        <div className="note">
          ※ 상단 1행은 그룹(단일컬럼 / 세트명)으로 병합된 헤더, 2행은 실제 세부 컬럼입니다. 데이터 행은 병합 시 3행부터 채워집니다.
        </div>
      </div>
    );
  };

  /** 스타일 */
  const styles = `
  .wrap { font-family: system-ui, -apple-system, Segoe UI, Roboto, 'Noto Sans KR', sans-serif; padding: 16px; }
  .topbar { display:flex; align-items:center; gap:12px; margin-bottom:16px; }
  .grid { display:grid; grid-template-columns: 1fr 1fr; gap:16px; min-height: 360px; }
  .pane { border:1px solid #e5e7eb; border-radius:12px; padding:12px; display:flex; flex-direction:column; }
  .pane h3 { margin:4px 0 10px; font-size:16px; }
  .pane .toolbar { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
  .btn { padding:6px 10px; font-size:13px; border-radius:8px; border:1px solid #d1d5db; background:#fff; cursor:pointer; }
  .btn.primary { background:#111827; color:#fff; border-color:#111827; }
  .btn.danger { background:#b91c1c; color:#fff; border-color:#b91c1c; }
  .list { border:1px solid #e5e7eb; border-radius:8px; padding:8px; overflow:auto; min-height:280px; background:#fafafa; }
  .row { display:flex; align-items:center; gap:8px; padding:6px 8px; border-bottom:1px dashed #eee; }
  .row:last-child { border-bottom:none; }
  .row label { flex:1; cursor:pointer; }
  .accordion { border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; }
  .setHeader { display:flex; align-items:center; gap:8px; padding:10px; background:#f3f4f6; border-bottom:1px solid #e5e7eb; cursor:pointer; }
  .setHeader .name { font-weight:600; }
  .setBody { padding:10px; background:#fff; }
  .chips { display:flex; flex-wrap:wrap; gap:8px; }
  .chip { display:inline-flex; align-items:center; gap:6px; padding:6px 8px; border-radius:999px; background:#eef2ff; border:1px solid #c7d2fe; font-size:12px; }
  .chip input { transform: translateY(1px); }
  .preview { margin-top:16px; border:1px solid #e5e7eb; border-radius:12px; padding:12px; }
  .preview.empty { color:#6b7280; font-size:14px; }
  .merged-table { border-collapse: collapse; width:100%; }
  .merged-table th { border:1px solid #e5e7eb; padding:8px; text-align:center; font-weight:600; }
  .lvl1 { background:#f9fafb; }
  .lvl2 { background:#fff; font-weight:500; }
  .note { color:#6b7280; font-size:12px; margin-top:8px; }
  .muted { color:#6b7280; font-size:12px; }
  .tip { margin: 8px 0 10px; padding:10px 12px; border-radius:10px; background:#f0fdf4; border:1px solid #bbf7d0; color:#166534; font-size:13px; }
  `;

  return (
    <div className="wrap">
      <style>{styles}</style>

      <div className="topbar">
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFile} />
        <button
          className="btn"
          onClick={() => {
            if (fileInputRef.current) fileInputRef.current.value = "";
            setOrigHeaders([]);
            setOrigRows([]);
            setOrig({ singles: [], sets: {} });
            setSingles([]);
            setSets({});
            setSingleSelection(new Set());
            setOpenSetNames(new Set());
            setSetSelection(new Set());
            setFieldSelection({});
          }}
        >
          파일/상태 초기화
        </button>
        <div className="muted">첫 시트의 1행을 헤더로 인식합니다. 업로드 후 좌/우에서 각각 항목을 정리하세요.</div>
      </div>

      <div className="grid">
        {/* 좌측: 단일 컬럼 */}
        <div className="pane">
          <h3>단일 컬럼 ({singles.length}개)</h3>
          <div className="tip">수험번호, 이름, 지원분야 역할만 고르시면 됩니다</div>
          <div className="toolbar">
            <button className="btn danger" onClick={deleteSingles_SelectedOnly}>선택만 삭제</button>
            <button className="btn danger" onClick={deleteSingles_ExceptSelected}>선택 제외 삭제</button>
            <button className="btn" onClick={resetSingles}>되돌리기</button>
          </div>

          <div className="list">
            {singles.length === 0 ? (
              <div className="muted">항목이 없습니다.</div>
            ) : (
              singles.map((s) => (
                <div key={s} className="row">
                  <input
                    type="checkbox"
                    checked={singleSelection.has(s)}
                    onChange={() => toggleSingle(s)}
                    id={`single-${s}`}
                  />
                  <label htmlFor={`single-${s}`}>{s}</label>
                </div>
              ))
            )}
          </div>
        </div>

        {/* 우측: 세트 컬럼 */}
        <div className="pane">
          <h3>세트 컬럼 ({Object.keys(sets).length}종)</h3>
          <div className="toolbar">
            <button className="btn danger" onClick={deleteSets_SelectedOnly}>선택 세트 삭제</button>
            <button className="btn danger" onClick={deleteSets_ExceptSelected}>선택 제외 세트 삭제</button>
            <button className="btn" onClick={deleteSelectedFieldsInSets}>선택 필드 삭제</button>
            <button className="btn" onClick={resetSets}>되돌리기</button>
          </div>

          <div className="accordion">
            {Object.entries(sets).length === 0 ? (
              <div className="row"><span className="muted">세트가 없습니다.</span></div>
            ) : (
              Object.entries(sets).map(([name, g]) => {
                const open = openSetNames.has(name);
                const selectedFields = fieldSelection[name] || new Set();
                const totalSpan = g.fields.length * g.indices.length;

                return (
                  <div key={name}>
                    <div className="setHeader" onClick={() => toggleSetOpen(name)}>
                      <input
                        type="checkbox"
                        checked={setSelection.has(name)}
                        onChange={(e) => { e.stopPropagation(); toggleSetSelected(name); }}
                        onClick={(e) => e.stopPropagation()}
                        title="세트 전체 선택(삭제용)"
                      />
                      <span className="name">{name}</span>
                      <span className="muted">| 인덱스 {g.indices.length}개 × 필드 {g.fields.length}개 = {totalSpan}</span>
                    </div>
                    {open && (
                      <div className="setBody">
                        <div className="muted" style={{ marginBottom: 8 }}>
                          ✔ 필드 선택 후 “선택 필드 삭제”를 누르면 해당 세트의 모든 인덱스에서 해당 필드들이 제거됩니다.
                        </div>
                        <div className="chips">
                          {g.fields.map((f) => (
                            <label key={f} className="chip">
                              <input
                                type="checkbox"
                                checked={selectedFields.has(f)}
                                onChange={() => toggleFieldSelected(name, f)}
                              />
                              <span>{f}</span>
                            </label>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })
            )}
          </div>
        </div>
      </div>

      {/* 병합 미리보기 + 병합 실행 */}
      <div style={{ marginTop: 16 }}>
        <h3>병합 미리보기</h3>
        <MergedPreview />
        <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
          <button className="btn primary" onClick={handleExportMerged}>
            병합하기 (엑셀 다운로드)
          </button>
        </div>
      </div>
    </div>
  );
}
