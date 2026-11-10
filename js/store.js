// 데이터 저장소: 기기(localStorage)에 먼저 저장하고, 로그인하면 sync.js가 클라우드와 동기화합니다.
const KEY = 'jywork-data-v1';
const SETTINGS_KEY = 'jywork-settings-v1';
export const KINDS = ['events', 'meetings', 'tasks'];

const empty = () => ({ events: [], meetings: [], tasks: [], tombs: {} });

export const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
export const nowIso = () => new Date().toISOString();

export const todayStr = (d = new Date()) => {
  const z = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
};

let data = load();
const listeners = new Set(); // 화면 갱신용
const changeListeners = new Set(); // 동기화용: [{op:'put'|'del', kind, id, item}]

function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return empty();
    const parsed = JSON.parse(raw);
    return { ...empty(), ...parsed, tombs: parsed.tombs || {} };
  } catch {
    return empty();
  }
}

// 항목별 스냅샷을 비교해서 무엇이 바뀌었는지 계산
function buildSnap(d) {
  const m = new Map();
  for (const k of KINDS) for (const x of d[k]) m.set(`${k}:${x.id}`, JSON.stringify(x));
  return m;
}
let snap = buildSnap(data);

function persist({ silent = false } = {}) {
  const next = buildSnap(data);
  const changes = [];
  if (!silent) {
    for (const [key, v] of next) {
      if (snap.get(key) !== v) {
        const [kind, id] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)];
        changes.push({ op: 'put', kind, id, item: JSON.parse(v) });
        delete data.tombs[key];
      }
    }
    const t = nowIso();
    for (const key of snap.keys()) {
      if (!next.has(key)) {
        const [kind, id] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)];
        changes.push({ op: 'del', kind, id, at: t });
        data.tombs[key] = t; // 삭제 기록(다른 기기에서도 지워지도록)
      }
    }
    // 90일 지난 삭제 기록 정리
    const old = new Date(Date.now() - 90 * 864e5).toISOString();
    for (const [k, v] of Object.entries(data.tombs)) if (v < old) delete data.tombs[k];
  }
  snap = next;
  try {
    localStorage.setItem(KEY, JSON.stringify(data));
  } catch {
    alert('저장 공간이 부족하거나 저장할 수 없습니다. 설정에서 백업 후 정리해주세요.');
  }
  if (changes.length) changeListeners.forEach((fn) => fn(changes));
  listeners.forEach((fn) => fn(data));
}

export const store = {
  get: () => data,
  subscribe: (fn) => listeners.add(fn),
  onChange: (fn) => changeListeners.add(fn),

  upsert(kind, item) {
    const list = data[kind];
    const now = nowIso();
    if (item.id) {
      const i = list.findIndex((x) => x.id === item.id);
      if (i >= 0) {
        list[i] = { ...list[i], ...item, updatedAt: now };
        persist();
        return list[i];
      }
    }
    const created = { ...item, id: item.id || uid(), createdAt: now, updatedAt: now };
    list.push(created);
    persist();
    return created;
  },

  remove(kind, id) {
    data[kind] = data[kind].filter((x) => x.id !== id);
    if (kind === 'meetings') data.tasks = data.tasks.filter((t) => t.meetingId !== id);
    persist();
  },

  find: (kind, id) => data[kind].find((x) => x.id === id),

  // 회의 요약에서 나온 할일을 할일 목록과 동기화
  syncMeetingTasks(meeting) {
    const keep = data.tasks.filter((t) => t.meetingId !== meeting.id);
    const old = data.tasks.filter((t) => t.meetingId === meeting.id);
    const now = nowIso();
    const fresh = (meeting.actions || []).map((a) => {
      const prev = old.find((o) => o.text === a.text);
      if (prev && prev.owner === (a.owner || '') && prev.due === (a.due || '') && prev.source === meeting.title) return prev;
      return {
        id: prev?.id || uid(),
        text: a.text,
        owner: a.owner || '',
        due: a.due || '',
        done: prev?.done || false,
        meetingId: meeting.id,
        source: meeting.title,
        createdAt: prev?.createdAt || now,
        updatedAt: now,
      };
    });
    data.tasks = [...keep, ...fresh];
    persist();
  },

  replaceAll(next) {
    data = { ...empty(), ...next, tombs: data.tombs };
    persist();
  },

  merge(next) {
    const now = nowIso();
    for (const kind of KINDS) {
      for (const raw of next[kind] || []) {
        const item = { ...raw, id: raw.id || uid(), updatedAt: now };
        const i = data[kind].findIndex((x) => x.id === item.id);
        if (i >= 0) data[kind][i] = { ...data[kind][i], ...item };
        else data[kind].push({ createdAt: now, ...item });
      }
    }
    persist();
  },

  // 다른 기기에서 받은 변경을 반영 (다시 업로드하지 않음)
  applyRemote(list) {
    for (const r of list) {
      const arr = data[r.kind];
      if (!arr) continue;
      const i = arr.findIndex((x) => x.id === r.id);
      if (r.deleted) {
        if (i >= 0) arr.splice(i, 1);
        data.tombs[`${r.kind}:${r.id}`] = r.updatedAt;
      } else if (i >= 0) arr[i] = r.data;
      else arr.push(r.data);
    }
    persist({ silent: true });
  },
};

export const settings = {
  get() {
    try {
      return { model: 'claude-opus-5-5', apiKey: '', firebase: null, ...JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}') };
    } catch {
      return { model: 'claude-opus-5-5', apiKey: '', firebase: null };
    }
  },
  set(v) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...settings.get(), ...v }));
  },
};
