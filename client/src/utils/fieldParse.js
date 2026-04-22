// 지원분야 텍스트에서 자주 쓰이는 구분자 후보
const CANDIDATES = ['>', '→', '›', '▶', '|', '/', ':', '-', '_', '·', '•'];

// 주어진 값 배열에서 각 구분자가 몇 개 값에 포함되는지 집계 (count > 0만 반환, 빈도 내림차순)
export function detectDelimiters(values) {
  const list = CANDIDATES.map((d) => ({
    delimiter: d,
    count: values.filter((v) => String(v ?? '').includes(d)).length
  })).filter((x) => x.count > 0);
  list.sort((a, b) => b.count - a.count || CANDIDATES.indexOf(a.delimiter) - CANDIDATES.indexOf(b.delimiter));
  return list;
}

// mode: 'first' = 첫 구분자 이전, 'last' = 마지막 구분자 이후
export function parseField(value, delimiter, mode = 'last') {
  const s = String(value ?? '');
  if (!delimiter || !s.includes(delimiter)) return s.trim();
  const parts = s.split(delimiter);
  return (mode === 'first' ? parts[0] : parts[parts.length - 1]).trim();
}
