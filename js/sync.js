// 탭·PC 동기화 (Firebase: 구글 로그인 + Firestore)
// - 기기에 먼저 저장하고(오프라인 가능), 로그인되어 있으면 클라우드와 주고받습니다.
// - 같은 항목을 두 기기에서 고치면 더 나중에 고친 쪽이 남습니다.
// - 데이터 경로: users/{내 계정}/items, users/{내 계정}/ink  (보안 규칙으로 본인만 접근)
import { store, settings, nowIso } from './store.js';
import { inkdb } from './inkdb.js';
import defaultConfig from './firebase-config.js';

const SDK = './vendor/firebase.mjs';
let fb = null; // SDK 함수들
let auth = null, db = null, user = null;
let unsub = null;
const listeners = new Set();
const state = { status: 'off', email: '', lastSync: '', error: '', pending: 0 };

function emit(patch) {
  Object.assign(state, patch);
  listeners.forEach((fn) => fn({ ...state }));
}

export function parseConfig(text) {
  if (!text) return null;
  if (typeof text === 'object') return text;
  try { return JSON.parse(text); } catch { /* 콘솔 코드 형식 */ }
  const cfg = {};
  for (const m of String(text).matchAll(/(\w+)\s*:\s*["']([^"']+)["']/g)) cfg[m[1]] = m[2];
  return cfg.apiKey && cfg.projectId ? cfg : null;
}

export const getConfig = () => parseConfig(settings.get().firebase) || defaultConfig || null;

const itemsCol = () => fb.collection(db, 'users', user.uid, 'items');
const inkCol = () => fb.collection(db, 'users', user.uid, 'ink');
const docId = (kind, id) => `${kind}_${id}`;
const clean = (o) => JSON.parse(JSON.stringify(o)); // undefined 제거

async function track(p) {
  emit({ pending: state.pending + 1, status: user ? 'syncing' : state.status });
  try {
    await p;
    emit({ pending: state.pending - 1, lastSync: nowIso(), error: '' });
  } catch (e) {
    emit({ pending: state.pending - 1, error: errMsg(e) });
  }
  if (user && !state.pending) emit({ status: state.error ? 'error' : 'synced' });
}

function errMsg(e) {
  const c = e?.code || '';
  if (c.includes('permission-denied')) return 'Firestore 보안 규칙을 확인해주세요 (권한 거부).';
  if (c.includes('unauthorized-domain')) return '이 주소가 Firebase 승인된 도메인에 등록되지 않았습니다.';
  if (c.includes('popup-closed')) return '로그인 창이 닫혔습니다.';
  if (c.includes('network')) return '인터넷 연결을 확인해주세요.';
  return e?.message || String(e);
}

// 여러 건을 500개 단위로 나눠 쓰기
async function writeMany(ops) {
  for (let i = 0; i < ops.length; i += 400) {
    const batch = fb.writeBatch(db);
    for (const op of ops.slice(i, i + 400)) {
      if (op.delete) batch.delete(op.ref);
      else batch.set(op.ref, op.data);
    }
    await batch.commit();
  }
}

const putOp = (kind, item) => ({
  ref: fb.doc(itemsCol(), docId(kind, item.id)),
  data: { kind, id: item.id, deleted: false, updatedAt: item.updatedAt || nowIso(), data: clean(item) },
});
const delOp = (kind, id, at) => ({
  ref: fb.doc(itemsCol(), docId(kind, id)),
  data: { kind, id, deleted: true, updatedAt: at || nowIso() },
});

/* ---------- 로그인 후 첫 동기화: 양쪽을 비교해서 최신으로 맞춤 ---------- */
async function initialMerge() {
  const remote = await fb.getDocs(itemsCol());
  const data = store.get();
  const toLocal = [];
  const toRemote = [];
  const seen = new Set();
  remote.forEach((d) => {
    const r = d.data();
    const key = `${r.kind}:${r.id}`;
    seen.add(key);
    const local = data[r.kind]?.find((x) => x.id === r.id);
    const lu = local?.updatedAt || '';
    const tomb = data.tombs[key];
    if (r.deleted) {
      if (local && lu <= r.updatedAt) toLocal.push(r);
      else if (local) toRemote.push(putOp(r.kind, local));
    } else if (tomb && tomb >= r.updatedAt) {
      toRemote.push(delOp(r.kind, r.id, tomb));
    } else if (!local || lu < r.updatedAt) {
      toLocal.push(r);
    } else if (lu > r.updatedAt) {
      toRemote.push(putOp(r.kind, local));
    }
  });
  for (const kind of ['events', 'meetings', 'tasks']) {
    for (const x of data[kind]) if (!seen.has(`${kind}:${x.id}`)) toRemote.push(putOp(kind, x));
  }
  for (const [key, at] of Object.entries(data.tombs)) {
    if (!seen.has(key)) {
      const [kind, id] = [key.slice(0, key.indexOf(':')), key.slice(key.indexOf(':') + 1)];
      toRemote.push(delOp(kind, id, at));
    }
  }
  if (toLocal.length) await applyRemoteItems(toLocal);
  await writeMany(toRemote);
  await uploadPendingInk();
  return { down: toLocal.length, up: toRemote.length };
}

async function applyRemoteItems(list) {
  store.applyRemote(list);
  // 다른 기기에서 지운 회의의 필기도 삭제
  for (const r of list) if (r.kind === 'meetings' && r.deleted) await inkdb.del(r.id).catch(() => {});
}

/* ---------- 실시간: 다른 기기에서 바뀐 내용 받기 ---------- */
function listen() {
  unsub?.();
  unsub = fb.onSnapshot(itemsCol(), (snap) => {
    const data = store.get();
    const list = [];
    snap.docChanges().forEach((ch) => {
      if (ch.type === 'removed' || ch.doc.metadata.hasPendingWrites) return;
      const r = ch.doc.data();
      const local = data[r.kind]?.find((x) => x.id === r.id);
      const lu = local?.updatedAt || '';
      if (r.deleted ? local && lu <= r.updatedAt : lu < r.updatedAt) list.push(r);
    });
    if (list.length) {
      applyRemoteItems(list);
      emit({ lastSync: nowIso() });
    }
  }, (e) => emit({ status: 'error', error: errMsg(e) }));
}

/* ---------- 필기(펜) 동기화: 쪽 단위로 저장 ---------- */
export async function pushInk(mid, ink) {
  if (!user) return;
  const pages = ink?.pages || [];
  const ops = pages.map((p, i) => ({
    ref: fb.doc(inkCol(), `${mid}_${i}`),
    data: { mid, i, n: pages.length, rev: ink.rev || '', ls: p.ls || 0, strokes: JSON.stringify(p.strokes) },
  }));
  const existing = await fb.getDocs(fb.query(inkCol(), fb.where('mid', '==', mid)));
  existing.forEach((d) => { if (d.data().i >= pages.length) ops.push({ ref: d.ref, delete: true }); });
  await writeMany(ops);
  await inkdb.put(mid, { ...ink, synced: true });
}

export async function fetchInk(mid) {
  if (!user) return null;
  const snap = await fb.getDocs(fb.query(inkCol(), fb.where('mid', '==', mid)));
  if (snap.empty) return null;
  const pages = [];
  let rev = '';
  snap.forEach((d) => {
    const x = d.data();
    pages[x.i] = { strokes: JSON.parse(x.strokes || '[]'), ls: x.ls || 0 };
    rev = x.rev;
  });
  const ink = { pages: pages.filter(Boolean), rev, synced: true };
  await inkdb.put(mid, ink);
  return ink;
}

async function uploadPendingInk() {
  const all = await inkdb.all().catch(() => ({}));
  for (const [mid, ink] of Object.entries(all)) {
    if (ink && !ink.synced && store.find('meetings', mid)) await pushInk(mid, ink);
  }
}

// 회의를 열 때: 기기에 없는/오래된 필기는 클라우드에서 받아옴
export async function loadInk(meeting) {
  const local = await inkdb.get(meeting.id).catch(() => null);
  if (meeting.inkPages && user && (!local || (local.rev || '') < (meeting.inkRev || ''))) {
    try { return (await fetchInk(meeting.id)) || local; } catch { return local; }
  }
  return local;
}

// 자동 저장 때마다 올리면 데이터가 많아서, 필기는 기기에 바로 저장하고 클라우드에는 몇 초 모아서 올립니다.
const inkPending = new Map(); // mid → { timer, ink }
export async function saveInk(mid, ink, { immediate = false } = {}) {
  const p = inkPending.get(mid);
  if (p) { clearTimeout(p.timer); inkPending.delete(mid); }
  if (!ink) {
    await inkdb.del(mid);
    if (user) track(fb.getDocs(fb.query(inkCol(), fb.where('mid', '==', mid))).then((s) => writeMany(s.docs.map((d) => ({ ref: d.ref, delete: true })))));
    return;
  }
  await inkdb.put(mid, { ...ink, synced: false });
  if (!user) return;
  const run = () => { inkPending.delete(mid); track(pushInk(mid, ink)); };
  if (immediate) run();
  else inkPending.set(mid, { ink, timer: setTimeout(run, 4000) });
}
export function flushInk() {
  for (const [mid, p] of [...inkPending]) {
    clearTimeout(p.timer);
    inkPending.delete(mid);
    if (user) track(pushInk(mid, p.ink));
  }
}

/* ---------- 시작 · 로그인 ---------- */
store.onChange((changes) => {
  if (!user) return;
  const ops = changes.map((c) => (c.op === 'put' ? putOp(c.kind, c.item) : delOp(c.kind, c.id, c.at)));
  track(writeMany(ops));
});

export const sync = {
  state: () => ({ ...state }),
  subscribe: (fn) => { listeners.add(fn); fn({ ...state }); },
  signedIn: () => !!user,
  configured: () => !!getConfig(),

  async init() {
    const cfg = getConfig();
    if (!cfg) { emit({ status: 'off' }); return; }
    try {
      fb = await import(SDK);
      const app = fb.initializeApp(cfg, 'jywork');
      auth = fb.getAuth(app);
      const cache = fb.persistentLocalCache({ tabManager: fb.persistentMultipleTabManager() });
      db = fb.initializeFirestore(app, cfg.emulator ? {} : { localCache: cache });
      if (cfg.emulator) {
        fb.connectAuthEmulator(auth, `http://${cfg.emulator}:9099`, { disableWarnings: true });
        fb.connectFirestoreEmulator(db, cfg.emulator, 8080);
      }
      emit({ status: 'signedout' });
      fb.getRedirectResult(auth).catch((e) => emit({ error: errMsg(e) }));
      fb.onAuthStateChanged(auth, (u) => (u ? this.start(u) : this.stop()));
    } catch (e) {
      emit({ status: 'error', error: '동기화 초기화 실패: ' + errMsg(e) });
    }
  },

  async start(u) {
    user = u;
    emit({ status: 'syncing', email: u.email || '', error: '' });
    try {
      const r = await initialMerge();
      listen();
      emit({ status: 'synced', lastSync: nowIso() });
      return r;
    } catch (e) {
      emit({ status: 'error', error: errMsg(e) });
    }
  },

  stop() {
    unsub?.();
    unsub = null;
    user = null;
    emit({ status: auth ? 'signedout' : 'off', email: '', pending: 0 });
  },

  async signIn() {
    if (!auth) throw new Error('먼저 Firebase 연결 설정을 저장해주세요.');
    const provider = new fb.GoogleAuthProvider();
    provider.setCustomParameters({ prompt: 'select_account' });
    try {
      await fb.signInWithPopup(auth, provider);
    } catch (e) {
      if (/popup-blocked|operation-not-supported/.test(e?.code || '')) return fb.signInWithRedirect(auth, provider);
      throw new Error(errMsg(e));
    }
  },

  // 에뮬레이터 테스트 전용 로그인
  async signInTest(token) {
    if (!getConfig()?.emulator) throw new Error('emulator only');
    const { signInWithCredential } = fb;
    return signInWithCredential(auth, fb.GoogleAuthProvider.credential(token));
  },

  signOut: () => auth && fb.signOut(auth),
  syncNow: () => user && sync.start(user),
};
