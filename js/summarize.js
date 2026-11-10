// 오프라인 자동 요약: API 키 없이도 동작하는 규칙 기반 요약기
// - 핵심 문장 추출(단어 빈도 점수)
// - 결정사항: 결정/확정/합의/승인 등이 들어간 문장
// - 할일: 담당/까지/하기로/요청/TODO 등이 들어간 문장, @이름·날짜 자동 인식

const STOP = new Set(('그리고 그러나 하지만 또한 및 등 것 수 이 그 저 있다 있음 없음 했다 한다 하는 하고 ' +
  '대한 대해 위해 통해 관련 이번 다음 오늘 내일 우리 회의 부분 경우 정도 때문 the a an of to and').split(' '));

const DECISION_RE = /(결정|확정|합의|승인|채택|결론|정하기로|하기로 함|로 정함|동의)/;
const ACTION_RE = /(담당|까지|하기로|요청|예정|TODO|todo|할 일|할일|진행할|준비|검토 필요|작성 필요|공유 필요|@\S+|^\s*[-*]\s*\[\s?\])/;

export function splitSentences(text) {
  return text
    .replace(/\r/g, '')
    .split(/\n+|(?<=[.!?。])\s+|(?<=다\.)\s*|(?<=음\.)\s*|(?<=함\.)\s*/)
    .map((s) => s.replace(/^\s*([-*•·]|\d+[.)])\s*/, '').replace(/^\[\s?[xX ]?\]\s*/, '').trim())
    .filter((s) => s.length > 1);
}

function tokens(s) {
  return (s.toLowerCase().match(/[가-힣a-z0-9]{2,}/g) || [])
    .map((w) => w.replace(/(은|는|이|가|을|를|에|의|와|과|도|로|으로|에서|에게|까지|부터)$/, ''))
    .filter((w) => w.length > 1 && !STOP.has(w));
}

export function parseOwner(s) {
  const at = s.match(/@([가-힣A-Za-z0-9_]+)/);
  if (at) return at[1];
  const m = s.match(/담당\s*[:：]?\s*([가-힣A-Za-z]{2,10})/) || s.match(/([가-힣]{2,4})\s*(님|씨|과장|대리|팀장|부장|차장|주임|사원|매니저)/);
  return m ? m[1] + (m[2] && !/님|씨/.test(m[2]) ? ' ' + m[2] : '') : '';
}

export function parseDue(s, base = new Date()) {
  const y = base.getFullYear();
  const z = (n) => String(n).padStart(2, '0');
  let m = s.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (m) return `${m[1]}-${z(m[2])}-${z(m[3])}`;
  m = s.match(/(\d{1,2})\s*[/.월]\s*(\d{1,2})\s*일?/);
  if (m) return `${y}-${z(m[1])}-${z(m[2])}`;
  const add = (days) => {
    const d = new Date(base);
    d.setDate(d.getDate() + days);
    return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
  };
  if (/오늘/.test(s)) return add(0);
  if (/내일/.test(s)) return add(1);
  if (/모레/.test(s)) return add(2);
  if (/다음\s*주|차주/.test(s)) return add(7);
  if (/이번\s*주/.test(s)) return add((5 - base.getDay() + 7) % 7);
  return '';
}

export function localSummarize(text, baseDate) {
  const sents = splitSentences(text);
  if (!sents.length) return { summary: '', decisions: [], actions: [], keywords: [] };

  const freq = {};
  sents.forEach((s) => tokens(s).forEach((w) => (freq[w] = (freq[w] || 0) + 1)));
  const keywords = Object.entries(freq).filter(([w]) => !/^\d+$/.test(w) && !/\d{2,}/.test(w.slice(0, 2))).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([w]) => w);

  const scored = sents.map((s, i) => {
    const t = tokens(s);
    const score = t.reduce((acc, w) => acc + (freq[w] || 0), 0) / Math.sqrt(t.length + 1)
      + (DECISION_RE.test(s) ? 2 : 0) + (i === 0 ? 1 : 0);
    return { s, i, score };
  });
  const n = Math.min(4, sents.length, Math.max(2, Math.ceil(sents.length / 3)));
  const top = [...scored].sort((a, b) => b.score - a.score).slice(0, n).sort((a, b) => a.i - b.i);

  const decisions = sents.filter((s) => DECISION_RE.test(s));
  const base = baseDate ? new Date(baseDate + 'T00:00:00') : new Date();
  const actions = sents
    .filter((s) => ACTION_RE.test(s) && !DECISION_RE.test(s))
    .map((s) => ({ text: s, owner: parseOwner(s), due: parseDue(s, base) }));

  return {
    summary: top.map((x) => x.s).join(' '),
    decisions,
    actions,
    keywords,
  };
}
