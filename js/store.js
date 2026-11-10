// 데이터 저장소: 브라우저 localStorage에 모든 데이터를 보관합니다.
const KEY = 'jywork-data-v1';
const SETTINGS_KEY = 'jywork-settings-v1';

const empty = () => ({ events: [], meetings: [], tasks: [] });

export const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

export const todayStr = (d = new Date()) => {
  const z = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
};

let data = load();
const listeners = new Set();

function load() {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return empty();
    const parsed = JSON.parse(raw);
    return { ...empty(), ...parsed };
  } catch {
    return empty();
  }
}

function persist() {
  try {
    localStorage.setItem(KEY, JSON.stringify(data));
  } catch (e) {
    alert('저장 공간이 부족하거나 저장할 수 없습니다. 설정에서 백업 후 정리해주세요.');
  }
  listeners.forEach((fn) => fn(data));
}

export const store = {
  get: () => data,
  subscribe: (fn) => listeners.add(fn),

  upsert(kind, item) {
    const list = data[kind];
    const now = new Date().toISOString();
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
    const fresh = (meeting.actions || []).map((a) => {
      const prev = old.find((o) => o.text === a.text);
      return {
        id: prev?.id || uid(),
        text: a.text,
        owner: a.owner || '',
        due: a.due || '',
        done: prev?.done || false,
        meetingId: meeting.id,
        source: meeting.title,
        createdAt: prev?.createdAt || new Date().toISOString(),
      };
    });
    data.tasks = [...keep, ...fresh];
    persist();
  },

  replaceAll(next) {
    data = { ...empty(), ...next };
    persist();
  },

  merge(next) {
    for (const kind of ['events', 'meetings', 'tasks']) {
      for (const item of next[kind] || []) {
        const i = data[kind].findIndex((x) => x.id === item.id);
        if (i >= 0) data[kind][i] = item;
        else data[kind].push({ ...item, id: item.id || uid() });
      }
    }
    persist();
  },
};

export const settings = {
  get() {
    try {
      return { model: 'claude-opus-5-5', apiKey: '', ...JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}') };
    } catch {
      return { model: 'claude-opus-5-5', apiKey: '' };
    }
  },
  set(v) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...settings.get(), ...v }));
  },
};
