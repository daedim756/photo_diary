import Anthropic from "https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@1.8.0/+esm";
import * as od from './onedrive.js';

const $ = s => document.querySelector(s);
const EXTRA_COUNT = 4, SLOT_COUNT = 1 + EXTRA_COUNT;
const SEARCH_RADIUS = 300; // m

/* ================= 알림 ================= */
let toastTimer;
function toast(msg, ms = 3500) {
  const t = $('#toast'); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); if (ms) toastTimer = setTimeout(() => (t.hidden = true), ms);
}
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* ================= 설정 (이 기기에만 저장) ================= */
const settings = {
  get kakao() { try { return localStorage.getItem('kakaoKey') || ''; } catch { return ''; } },
  get claude() { try { return localStorage.getItem('claudeKey') || ''; } catch { return ''; } },
  save(kakao, claude) { localStorage.setItem('kakaoKey', kakao); localStorage.setItem('claudeKey', claude); },
};
// 카카오 키는 원드라이브에도 저장해서 폰에서 다시 입력하지 않게 함 (Claude 키는 비밀이라 기기에만)
async function pushSettings() {
  if (!od.signedIn()) return;
  try { await od.putFile('settings.json', JSON.stringify({ kakaoKey: settings.kakao }), 'application/json'); } catch { /* 다음 동기화 때 다시 */ }
}
async function pullSettings() {
  const blob = await od.download('settings.json').catch(() => null);
  if (!blob) { if (settings.kakao) pushSettings(); return; }
  try {
    const s = JSON.parse(await blob.text());
    if (s.kakaoKey && s.kakaoKey !== settings.kakao) { localStorage.setItem('kakaoKey', s.kakaoKey); updateKeyNote(); }
  } catch { /* 무시 */ }
}
function updateKeyNote() {
  const missing = [];
  if (!settings.kakao) missing.push('카카오 키(주변 가게 찾기)');
  if (!settings.claude) missing.push('Claude 키(사진 판별)');
  const n = $('#keyNote');
  n.hidden = !missing.length;
  n.textContent = missing.length ? `설정에서 ${missing.join(', ')}를 넣으면 장소 추천을 쓸 수 있어요.` : '';
}
$('#originText').textContent = location.origin;
$('#btnSettings').onclick = () => {
  $('#kakaoKey').value = settings.kakao; $('#claudeKey').value = settings.claude;
  $('#settings').hidden = false;
};
$('#btnSettingsClose').onclick = () => ($('#settings').hidden = true);
$('#btnSettingsSave').onclick = () => {
  const oldKakao = settings.kakao;
  settings.save($('#kakaoKey').value.trim(), $('#claudeKey').value.trim());
  $('#settings').hidden = true; updateKeyNote();
  toast('저장했어요.');
  pushSettings();
  if (oldKakao && oldKakao !== settings.kakao) location.reload(); // 지도 SDK는 키를 바꾸면 새로 불러와야 함
};

/* ================= IndexedDB ================= */
const DB_NAME = 'photo-diary-app', STORE = 'entries';
function openDB() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: 'id' });
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
}
async function tx(mode, fn) {
  const db = await openDB();
  return new Promise((res, rej) => {
    const t = db.transaction(STORE, mode), req = fn(t.objectStore(STORE));
    t.oncomplete = () => res(req && req.result); t.onerror = () => rej(t.error);
  });
}
const dbAll = () => tx('readonly', s => s.getAll());
const dbPut = e => tx('readwrite', s => s.put(e));
const dbDel = id => tx('readwrite', s => s.delete(id));

/* ================= EXIF: 촬영일 + GPS (JPEG) ================= */
async function readExif(file) {
  try {
    const v = new DataView(await file.slice(0, 512 * 1024).arrayBuffer());
    if (v.getUint16(0) !== 0xFFD8) return {};
    let off = 2;
    while (off < v.byteLength - 4) {
      const marker = v.getUint16(off);
      if ((marker & 0xFF00) !== 0xFF00) return {};
      const len = v.getUint16(off + 2);
      if (marker === 0xFFE1 && v.getUint32(off + 4) === 0x45786966) return parseTiff(v, off + 10);
      if (marker === 0xFFDA) break;
      off += 2 + len;
    }
  } catch { /* 정보 없음 */ }
  return {};
}
function parseTiff(v, start) {
  const le = v.getUint16(start) === 0x4949;
  const u16 = o => v.getUint16(start + o, le), u32 = o => v.getUint32(start + o, le);
  const ascii = (o, n) => { let s = ''; for (let i = 0; i < n; i++) { const c = v.getUint8(start + o + i); if (!c) break; s += String.fromCharCode(c); } return s; };
  const ifd = o => {
    const map = {}, n = u16(o);
    for (let i = 0; i < n; i++) {
      const e = o + 2 + i * 12, tag = u16(e), type = u16(e + 2), cnt = u32(e + 4);
      if (type === 2) map[tag] = ascii(cnt > 4 ? u32(e + 8) : e + 8, cnt);
      else if (type === 4 || type === 13) map[tag] = u32(e + 8);
      else if (type === 5) {
        const p = u32(e + 8), arr = [];
        for (let k = 0; k < cnt; k++) { const d = u32(p + k * 8 + 4); arr.push(d ? u32(p + k * 8) / d : 0); }
        map[tag] = arr;
      }
    }
    return map;
  };
  const out = {};
  const ifd0 = ifd(u32(4));
  let dt = null;
  if (ifd0[0x8769]) {
    const ex = ifd(ifd0[0x8769]);
    dt = ex[0x9003] || ex[0x9004];
    const off = ex[0x9011] || ex[0x9010]; // 촬영 시각의 시간대 (예: +09:00)
    if (typeof off === 'string' && /^[+-]\d{2}:\d{2}$/.test(off)) out.offset = off;
  }
  dt = dt || ifd0[0x0132];
  const m = dt && dt.match(/^(\d{4}):(\d{2}):(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?/);
  if (m && m[1] !== '0000') {
    out.date = `${m[1]}-${m[2]}-${m[3]}`;
    if (m[4]) out.taken = `${out.date}T${m[4]}:${m[5]}:${m[6]}`; // 초 단위 촬영 시각 (원드라이브에서 같은 사진 찾을 때 씀)
  }
  if (ifd0[0x8825]) {
    const g = ifd(ifd0[0x8825]);
    const dms = a => Array.isArray(a) && a.length >= 3 ? a[0] + a[1] / 60 + a[2] / 3600 : null;
    let lat = dms(g[2]), lng = dms(g[4]);
    if (lat != null && lng != null && (lat || lng)) {
      if (g[1] === 'S') lat = -lat;
      if (g[3] === 'W') lng = -lng;
      out.lat = lat; out.lng = lng;
    }
  }
  return out;
}
const toISO = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

/* ================= 이미지 축소: 원본 용량의 약 1/15 ================= */
const SHRINK_RATIO = 15, MIN_TARGET = 40 * 1024;
async function loadImage(blob) {
  const url = URL.createObjectURL(blob);
  try { return await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = url; }); }
  finally { setTimeout(() => URL.revokeObjectURL(url), 0); }
}
async function encode(img, max, q) {
  const scale = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
  const c = document.createElement('canvas');
  c.width = Math.round(img.naturalWidth * scale); c.height = Math.round(img.naturalHeight * scale);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
  return new Promise(res => c.toBlob(res, 'image/jpeg', q));
}
async function shrink(file) {
  const img = await loadImage(file);
  const target = Math.max(file.size / SHRINK_RATIO, MIN_TARGET);
  let max = Math.min(1600, Math.max(img.naturalWidth, img.naturalHeight)), best = null;
  while (true) {
    for (const q of [0.82, 0.72, 0.62]) {
      best = await encode(img, max, q);
      if (best.size <= target) return best;
    }
    if (max <= 480) return best;
    max = Math.max(480, Math.round(max * 0.8));
  }
}
const fmtSize = n => n >= 1048576 ? (n / 1048576).toFixed(1) + 'MB' : Math.round(n / 1024) + 'KB';
const blobToBase64 = b => new Promise(r => { const f = new FileReader(); f.onload = () => r(String(f.result).split(',')[1]); f.readAsDataURL(b); });

/* ================= Claude: 사진 속 가게 종류 판별 ================= */
const CLASSIFY_SCHEMA = {
  type: 'object',
  properties: {
    subject: { type: 'string', description: '사진에 보이는 주된 대상, 한국어로 짧게 (예: 짬뽕, 아이스 아메리카노, 은행나무 공원)' },
    place_type: { type: 'string', description: '이 사진을 찍었을 법한 가게나 장소 종류, 한국어 (예: 중국집, 카페, 공원)' },
    keywords: { type: 'array', items: { type: 'string' }, description: '카카오맵에서 주변 장소를 찾을 검색어 1~3개, 짧은 한국어 (예: 중식, 짬뽕)' },
    category: { type: 'string', enum: ['FD6', 'CE7', 'AT4', 'CT1', 'OTHER'], description: '카카오 카테고리: FD6 음식점, CE7 카페, AT4 관광명소, CT1 문화시설, OTHER 그 외' },
  },
  required: ['subject', 'place_type', 'keywords', 'category'],
  additionalProperties: false,
};
async function classifyPhoto(blob) {
  const client = new Anthropic({ apiKey: settings.claude, dangerouslyAllowBrowser: true });
  const small = await encode(await loadImage(blob), 1000, 0.8);
  const res = await client.beta.messages.create({
    model: 'claude-opus-5-5',
    max_tokens: 4000,
    betas: ['server-side-fallback-2026-07-01'],
    fallbacks: 'default',
    output_config: { effort: 'low', format: { type: 'json_schema', schema: CLASSIFY_SCHEMA } },
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: await blobToBase64(small) } },
        { type: 'text', text: '한국에서 찍은 일상 사진이에요. 사진에 무엇이 보이는지, 그리고 어떤 종류의 가게나 장소에서 찍었을지 알려주세요. 검색어는 그 주변 가게를 지도에서 찾는 데 쓸 거예요.' },
      ],
    }],
  });
  if (res.stop_reason === 'refusal') throw new Error('사진을 판별하지 못했어요');
  const text = res.content.find(b => b.type === 'text');
  if (!text) throw new Error('판별 결과가 비어 있어요');
  return JSON.parse(text.text);
}

/* ================= 카카오: 주변 장소 검색 ================= */
let kakaoReady = null;
function loadKakao() {
  if (!settings.kakao) return Promise.reject(new Error('카카오 키가 없어요'));
  if (kakaoReady) return kakaoReady;
  kakaoReady = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${encodeURIComponent(settings.kakao)}&libraries=services&autoload=false`;
    s.onload = () => window.kakao && window.kakao.maps ? kakao.maps.load(() => res(kakao)) : rej(new Error('카카오 지도를 불러오지 못했어요'));
    s.onerror = () => rej(new Error(`카카오 지도를 불러오지 못했어요. 카카오 개발자 사이트의 JavaScript SDK 도메인에 ${location.origin} 이 등록돼 있는지 확인해주세요.`));
    document.head.appendChild(s);
  });
  kakaoReady.catch(() => { kakaoReady = null; });
  return kakaoReady;
}
function placesCall(method, arg, lat, lng, radius = SEARCH_RADIUS) {
  return new Promise(res => {
    const S = kakao.maps.services;
    new S.Places()[method](arg, (data, status) => res(status === S.Status.OK ? data : []), {
      location: new kakao.maps.LatLng(lat, lng), radius, sort: S.SortBy.DISTANCE,
    });
  });
}

// 후보 찾기 상태. 버튼은 후보를 "더하거나 앞으로 올리기"만 하고 절대 빼지 않음.
const CATS = [['FD6', '음식점'], ['CE7', '카페'], ['AT4', '관광명소'], ['CT1', '문화시설']];
const RADII = [SEARCH_RADIUS, 700, 1500];
let sg = null;
async function runSearches(s, onlyCats) {
  await loadKakao();
  const { lat, lng } = s.cur, r = RADII[s.radiusIdx];
  const add = (list, matched) => {
    for (const p of list) {
      const prev = s.found.get(p.id);
      if (!prev) s.found.set(p.id, { ...p, matched });
      else if (matched) prev.matched = true;
    }
  };
  if (!onlyCats && s.kind) for (const kw of (s.kind.keywords || []).slice(0, 3)) add(await placesCall('keywordSearch', kw, lat, lng, r), true);
  if (!onlyCats) for (const w of s.words || []) add(await placesCall('keywordSearch', w, lat, lng, r), false);
  for (const c of onlyCats || [...s.cats]) add(await placesCall('categorySearch', c, lat, lng, r), false);
}
const fmtDist = d => (d === '' || d == null ? '' : +d >= 1000 ? `${(+d / 1000).toFixed(1)}km` : `${d}m`);
const sortedCandidates = s => s.mode === 'search' ? [...s.found.values()] : [...s.found.values()].sort((a, b) =>
  (b.matched - a.matched)
  || ((b.category_group_code === s.priority) - (a.category_group_code === s.priority))
  || (+a.distance - +b.distance));

/* ================= 편집기 ================= */
let draft = null, pickTarget = 0, map = null, markers = [];
const editor = $('#editor'), pick = $('#pick');
const slotLabels = ['대표 사진<br>클릭 또는 끌어놓기', '추가 사진 1', '추가 사진 2', '추가 사진 3', '추가 사진 4'];

function openEditor(entry) {
  const photos = entry ? [entry.main, ...(entry.extras || [])] : [];
  while (photos.length < SLOT_COUNT) photos.push(null);
  const names = entry && entry.photoNames ? [...entry.photoNames] : [];
  while (names.length < SLOT_COUNT) names.push(null);
  draft = {
    id: entry ? entry.id : (crypto.randomUUID ? crypto.randomUUID() : String(Date.now())),
    created: entry ? entry.created : Date.now(),
    photos: photos.slice(0, SLOT_COUNT), sizes: [], names: names.slice(0, SLOT_COUNT), meta: [],
    lat: entry?.lat ?? null, lng: entry?.lng ?? null, place: entry?.placeInfo || null,
    taken: entry?.takenAt ?? null,
  };
  $('#fDate').value = entry ? entry.date : '';
  $('#fPlace').value = entry ? entry.place || '' : '';
  $('#fMemo').value = entry ? entry.memo || '' : '';
  setHint(entry ? '' : '대표 사진을 올리면 촬영 날짜가 자동으로 들어갑니다.');
  clearSuggestions();
  setLocText();
  $('#btnRetryLoc').hidden = true;
  $('#placeSearch').hidden = true;
  stopMapPick();
  renderSlots();
  if (view !== 'list') showView('list');
  editor.hidden = false;
  editor.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
function closeEditor() { editor.hidden = true; draft = null; }
function setHint(text, ok) { const h = $('#dateHint'); h.textContent = text; h.className = 'hint' + (ok ? ' ok' : ''); }
function setLocText(msg) {
  $('#locText').textContent = msg || (draft && draft.lat != null
    ? `위치: ${draft.lat.toFixed(5)}, ${draft.lng.toFixed(5)}`
    : '위치: 대표 사진을 올리면 찍은 위치를 읽어요.');
}
function setStatus(msg) { $('#suggestStatus').textContent = msg || ''; }
function clearSuggestions() { sg = null; $('#candidates').innerHTML = ''; $('#chips').innerHTML = ''; $('#map').hidden = true; setStatus(''); }

function renderSlots() {
  document.querySelectorAll('.slot').forEach(el => {
    const i = +el.dataset.slot, blob = draft.photos[i];
    if (el._url) { URL.revokeObjectURL(el._url); el._url = null; }
    if (blob) {
      el._url = URL.createObjectURL(blob);
      el.innerHTML = `<img src="${el._url}" alt="" draggable="false"><button class="x" title="빼기" aria-label="사진 빼기">✕</button>`
        + (i > 0 ? '<button class="mk-main" title="이 사진을 대표 사진으로">대표로</button>' : '')
        + (draft.sizes[i] ? `<span class="size">${draft.sizes[i]}</span>` : '');
      el.classList.add('filled');
      el.draggable = true;
    } else { el.innerHTML = slotLabels[i]; el.classList.remove('filled'); el.draggable = false; }
  });
}

// 두 칸의 사진 자리 바꾸기 (복제 아님). 대표 사진이 바뀌면 그 사진의 날짜·위치를 씀.
function swapSlots(a, b) {
  if (!draft || a === b) return;
  for (const k of ['photos', 'names', 'sizes', 'meta']) {
    const arr = draft[k];
    [arr[a], arr[b]] = [arr[b], arr[a]];
  }
  renderSlots();
  if (a !== 0 && b !== 0) return;
  const m = draft.meta[0];
  if (m && m.date) { $('#fDate').value = m.date; setHint('대표 사진의 촬영 날짜로 바꿨어요.', true); }
  if (m) { draft.taken = m.taken || null; draft.offset = m.offset || null; }
  if (m && m.lat != null) { draft.lat = m.lat; draft.lng = m.lng; setLocText(); suggest(); }
  toast('대표 사진을 바꿨어요.');
}

async function setPhoto(i, file) {
  if (!draft || !file || !file.type.startsWith('image/')) return;
  const exif = await readExif(file); // 모든 칸에서 읽어 둠: 나중에 대표 사진으로 바꿀 때 씀
  if (i === 0) {
    draft.taken = exif.taken || null; draft.offset = exif.offset || null; // 찍은 시각: Week 탭에서 하루 동선 순서에 씀
    if (exif.date) { $('#fDate').value = exif.date; setHint('사진의 촬영 날짜를 읽었습니다.', true); }
    else { $('#fDate').value = toISO(new Date(file.lastModified)); setHint('사진에 촬영 날짜 정보가 없어 파일 날짜를 넣었습니다. 필요하면 고쳐주세요.'); }
  }
  try { draft.photos[i] = await shrink(file); }
  catch { toast('이 사진 형식은 열 수 없어요. 아이폰 HEIC 사진이면 JPG로 바꿔서 올려주세요.'); return; }
  draft.sizes[i] = `${fmtSize(file.size)} → ${fmtSize(draft.photos[i].size)}`;
  draft.names[i] = null; // 새 사진: 원드라이브에 새 이름으로 올림
  draft.meta[i] = exif;
  renderSlots();
  if (i === 0) {
    if (exif.lat != null) { draft.lat = exif.lat; draft.lng = exif.lng; setLocText(); suggest(); }
    else {
      draft.lat = draft.lng = null; clearSuggestions();
      draft.taken = exif.taken || null; draft.offset = exif.offset || null;
      lookupFromOneDrive();
    }
  }
}

// 갤러리에서 고르면 폰이 위치를 지우므로, 원드라이브 카메라 백업의 같은 사진(촬영 시각으로 찾음)에서 위치를 가져옴
async function lookupFromOneDrive() {
  const cur = draft;
  if (!cur) return;
  $('#btnRetryLoc').hidden = true;
  if (!cur.taken) { setLocText('위치: 이 사진에는 위치 정보가 없어요. 지금 그 장소라면 "현재 위치 사용"을 눌러주세요.'); return; }
  if (!od.configured() || !od.signedIn()) {
    setLocText('위치: 사진에 위치가 없어요. 원드라이브에 연결돼 있으면 카메라 백업에서 위치를 찾아올 수 있어요.');
    return;
  }
  setLocText('위치: 원드라이브 카메라 백업에서 같은 사진을 찾는 중…');
  try {
    const r = await od.findPhotoLocation(cur.taken, cur.offset);
    if (draft !== cur) return;
    let loc = r && r.lat != null ? r : null;
    if (r && r.blob) { const ex = await readExif(r.blob); if (ex.lat != null) loc = { lat: ex.lat, lng: ex.lng }; }
    if (draft !== cur) return;
    if (loc) {
      cur.lat = loc.lat; cur.lng = loc.lng;
      setLocText(`위치: ${loc.lat.toFixed(5)}, ${loc.lng.toFixed(5)} (원드라이브 백업에서 찾음)`);
      suggest();
    } else if (r) {
      setLocText('위치: 원드라이브에서 같은 사진을 찾았지만, 그 사진에도 위치 정보가 없어요. 카메라 앱의 위치 태그가 꺼져 있었을 수 있어요.');
    } else {
      setLocText('위치: 원드라이브 카메라 백업에서 같은 사진을 아직 못 찾았어요. 백업이 끝난 뒤 "다시 찾기"를 눌러주세요.');
      $('#btnRetryLoc').hidden = false;
    }
  } catch (e) {
    if (draft !== cur) return;
    if (e instanceof od.NeedsConsent || e instanceof od.AuthNeeded) {
      setSyncUI();
      setLocText('위치: 원드라이브 사진을 읽으려면 한 번 허락이 필요해요. 이 기록을 저장하거나 취소한 뒤, 위쪽 "원드라이브 다시 연결"을 눌러 허용해주세요.');
    } else {
      setLocText(`위치: 원드라이브에서 찾지 못했어요 (${e.message}).`);
      $('#btnRetryLoc').hidden = false;
    }
  }
}

$('#btnRetryLoc').onclick = () => lookupFromOneDrive();
$('#btnHere').onclick = () => {
  if (!navigator.geolocation) { toast('이 브라우저는 위치를 알려주지 않아요.'); return; }
  setLocText('위치: 현재 위치를 찾는 중…');
  navigator.geolocation.getCurrentPosition(
    p => { draft.lat = p.coords.latitude; draft.lng = p.coords.longitude; setLocText(); suggest(); },
    () => { setLocText(); toast('현재 위치를 가져오지 못했어요. 위치 권한을 확인해주세요.'); },
    { enableHighAccuracy: true, timeout: 10000 },
  );
};

let suggestSeq = 0;
async function suggest(opts = {}) {
  if (!draft || draft.lat == null) return;
  const seq = ++suggestSeq, cur = draft;
  clearSuggestions();
  if (!settings.kakao) { setStatus('주변 가게 찾기를 쓰려면 위쪽 "설정"에 카카오 키를 넣어주세요.'); return; }
  let kind = null;
  if (settings.claude && cur.photos[0]) {
    setStatus('사진을 보고 어떤 곳인지 알아보는 중…');
    try { kind = await classifyPhoto(cur.photos[0]); }
    catch (e) { setStatus(`사진 판별 실패(${e.status || e.message}). 주변 음식점·카페로 찾을게요.`); }
  }
  if (seq !== suggestSeq || draft !== cur) return;
  const s = { cur, kind, radiusIdx: 0, cats: new Set(['FD6', 'CE7']), priority: null, showAll: false, found: new Map(), pickedId: null };
  // 지도에서 찍은 곳이면 가게뿐 아니라 정류장·산·공원·명소·역도 함께 (가까운 순)
  if (opts.fromMap) { POI_CATS.forEach(c => s.cats.add(c)); s.words = POI_WORDS; }
  if (kind && kind.category !== 'OTHER') { s.cats.add(kind.category); s.priority = kind.category; }
  setStatus(kind ? `사진: ${kind.subject} → ${kind.place_type} 주변에서 찾는 중…` : '주변 장소를 찾는 중…');
  try {
    await runSearches(s);
    // 지도에서 찍은 곳: 이름 있는 장소가 없을 때(정류장, 길가 등)를 위해 그 지점의 주소도 맨 위에 넣음
    if (opts.fromMap) {
      const addr = await new Promise(res => new kakao.maps.services.Geocoder().coord2Address(cur.lng, cur.lat, (r, st) => {
        const a = st === kakao.maps.services.Status.OK && r[0];
        res(a ? (a.road_address ? a.road_address.address_name : a.address.address_name) : null);
      }));
      if (addr) s.found = new Map([['addr', { id: 'addr', place_name: addr, category_name: '이 지점의 주소', address_name: addr, distance: '0', x: String(cur.lng), y: String(cur.lat), matched: true }], ...s.found]);
    }
  }
  catch (e) { if (seq === suggestSeq) setStatus(e.message); return; }
  if (seq !== suggestSeq || draft !== cur) return;
  sg = s;
  renderCandidates();
}

function renderCandidates() {
  const s = sg;
  if (!s || draft !== s.cur) return;
  const all = sortedCandidates(s), r = RADII[s.radiusIdx];
  const shown = s.showAll ? all : all.slice(0, 8);
  if (s.mode === 'search') {
    setStatus(all.length ? `"${s.query}" 검색 결과 ${all.length}곳이에요. 맞는 곳을 눌러주세요.` : `"${s.query}"(으)로 찾은 곳이 없어요. 다른 이름으로 검색하거나 "지도에서 찍기"를 써보세요.`);
    $('#chips').innerHTML = '';
  } else {
    setStatus(all.length
      ? `${s.kind ? `사진: ${s.kind.subject} → ${s.kind.place_type}. ` : ''}반경 ${r}m 안에서 ${all.length}곳을 찾았어요.`
      : `반경 ${r}m 안에서 찾은 곳이 없어요. 범위를 넓히거나 장소를 직접 골라주세요.`);
    $('#chips').innerHTML = `<span class="chips-label">앞으로 올리기</span>`
      + CATS.map(([c, n]) => `<button type="button" class="chip${s.priority === c ? ' on' : ''}" data-cat="${c}">${n}</button>`).join('')
      + (s.radiusIdx < RADII.length - 1 ? `<button type="button" class="chip" data-wide="1">범위 넓히기 → ${RADII[s.radiusIdx + 1]}m</button>` : '');
  }
  const ol = $('#candidates');
  ol.innerHTML = shown.map((p, k) => `
    <li><button type="button" data-k="${k}"${s.pickedId === p.id ? ' class="picked"' : ''}>
      <span class="name">${esc(p.place_name)}${p.matched && s.kind ? `<span class="tag">${esc(s.kind.place_type)}</span>` : ''}</span>
      <span class="dist">${fmtDist(p.distance)}</span>
      <span class="meta">${esc((p.category_name || '').split(' > ').slice(-2).join(' · '))} · ${esc(p.road_address_name || p.address_name)}</span>
    </button></li>`).join('')
    + (all.length > shown.length ? `<li><button type="button" class="more" data-more="1">${all.length - shown.length}곳 더 보기</button></li>` : '');
  ol.querySelectorAll('button[data-k]').forEach(b => b.onclick = () => {
    const p = shown[+b.dataset.k];
    $('#fPlace').value = p.place_name;
    draft.place = { name: p.place_name, address: p.road_address_name || p.address_name, url: p.place_url, kakaoId: p.id };
    // 고른 가게의 좌표를 이 기록의 위치로 씀 (지도 탭에 정확히 표시됨)
    if (p.x && p.y) { draft.lat = +p.y; draft.lng = +p.x; setLocText(`위치: ${p.place_name} (${draft.lat.toFixed(5)}, ${draft.lng.toFixed(5)})`); }
    s.pickedId = p.id;
    ol.querySelectorAll('button[data-k]').forEach(x => x.classList.toggle('picked', x === b));
  });
  const more = ol.querySelector('[data-more]');
  if (more) more.onclick = () => { s.showAll = true; renderCandidates(); };
  drawMap(shown);
}

$('#chips').addEventListener('click', async ev => {
  const b = ev.target.closest('button'), s = sg;
  if (!b || !s) return;
  b.disabled = true;
  try {
    if (b.dataset.cat) {
      const c = b.dataset.cat;
      s.priority = s.priority === c ? null : c;
      if (!s.cats.has(c)) { s.cats.add(c); setStatus('찾는 중…'); await runSearches(s, [c]); }
    } else if (b.dataset.wide) {
      s.radiusIdx++; setStatus('범위를 넓혀서 찾는 중…'); await runSearches(s);
    }
  } catch (e) { setStatus(e.message); }
  if (sg === s) renderCandidates();
});

function drawMap(list) {
  const el = $('#map');
  if (!window.kakao || !kakao.maps || !draft) return;
  const hasLoc = draft.lat != null;
  if (!hasLoc && !list.length) return;
  const center = hasLoc ? new kakao.maps.LatLng(draft.lat, draft.lng) : new kakao.maps.LatLng(+list[0].y, +list[0].x);
  ensureMap(center, 3);
  markers.forEach(m => m.setMap(null)); markers = [];
  if (hasLoc) markers.push(new kakao.maps.Circle({ center, radius: 12, strokeWeight: 2, strokeColor: '#2e6bd6', fillColor: '#2e6bd6', fillOpacity: 0.6, map }));
  list.forEach(p => markers.push(new kakao.maps.Marker({ position: new kakao.maps.LatLng(+p.y, +p.x), title: p.place_name, map })));
  if (!hasLoc && list.length > 1) {
    const b = new kakao.maps.LatLngBounds();
    list.forEach(p => b.extend(new kakao.maps.LatLng(+p.y, +p.x)));
    map.setBounds(b, 30, 30, 30, 30);
  }
}

// 편집 화면의 작은 지도: 처음 한 번 만들고, "지도에서 찍기" 클릭도 여기서 받음
let pickingOnMap = false;
function ensureMap(center, level) {
  const el = $('#map');
  el.hidden = false;
  if (!map) {
    map = new kakao.maps.Map(el, { center, level });
    kakao.maps.event.addListener(map, 'click', ev => {
      if (!pickingOnMap || !draft) return;
      stopMapPick();
      const ll = ev.latLng;
      draft.lat = ll.getLat(); draft.lng = ll.getLng();
      setLocText(`위치: ${draft.lat.toFixed(5)}, ${draft.lng.toFixed(5)} (지도에서 고름)`);
      suggest({ fromMap: true });
    });
    // 지도를 움직이거나 확대하면 보이는 곳의 장소 이름을 다시 띄움
    kakao.maps.event.addListener(map, 'idle', () => {
      if (!pickingOnMap) return;
      clearTimeout(poiTimer); poiTimer = setTimeout(loadPois, 350);
    });
  } else { map.relayout(); map.setCenter(center); }
}

async function startMapPick() {
  if (!draft) return;
  if (!settings.kakao) { setStatus('지도를 쓰려면 위쪽 "설정"에 카카오 키를 넣어주세요.'); return; }
  try { await loadKakao(); } catch (e) { setStatus(e.message); return; }
  const last = current.find(e => e.lat != null);
  const center = draft.lat != null ? new kakao.maps.LatLng(draft.lat, draft.lng)
    : last ? new kakao.maps.LatLng(last.lat, last.lng) : new kakao.maps.LatLng(37.5665, 126.978);
  clearSuggestions();
  ensureMap(center, 4);
  markers.forEach(m => m.setMap(null)); markers = [];
  pickingOnMap = true;
  $('#map').classList.add('picking');
  map.relayout(); map.setCenter(center); // 지도 크기가 커졌으니 다시 맞춤
  setStatus('지도를 움직여서 간 곳을 찾아주세요.');
  loadPois();
}

/* ---------- 지도 위에 누를 수 있는 장소 이름 (산·정류장·명소 등) ---------- */
const POI_CATS = ['AT4', 'SW8', 'CT1']; // 관광명소, 지하철역, 문화시설
const POI_WORDS = ['산', '봉우리', '공원']; // 카카오 장소 검색에는 버스정류장이 없어서 제외
let poiOverlays = [], poiTimer = null, poiSeq = 0;
function clearPois() { poiOverlays.forEach(o => o.setMap(null)); poiOverlays = []; }
function stopMapPick() {
  pickingOnMap = false; poiSeq++; clearTimeout(poiTimer); clearPois();
  $('#map').classList.remove('picking');
}
async function loadPois() {
  if (!pickingOnMap || !map) return;
  const seq = ++poiSeq;
  if (map.getLevel() > 5) {
    clearPois();
    setStatus('지도를 조금 더 확대하면 산·정류장·명소 이름이 나와요. 이름을 누르거나, 아무 곳이나 눌러 고르세요.');
    return;
  }
  const S = kakao.maps.services, ps = new S.Places(), bounds = map.getBounds();
  const call = (method, arg) => new Promise(res => ps[method](arg, (d, st) => res(st === S.Status.OK ? d : []), { bounds }));
  const results = await Promise.all([...POI_CATS.map(c => call('categorySearch', c)), ...POI_WORDS.map(w => call('keywordSearch', w))]);
  if (seq !== poiSeq || !pickingOnMap) return;
  const found = new Map();
  results.flat().forEach(p => found.set(p.id, p));
  // 겹쳐서 누르기 힘들지 않게, 화면 가운데에서 가까운 곳부터 일부만 표시 (확대하면 더 보임)
  const c = map.getCenter(), cy = c.getLat(), cx = c.getLng();
  const shownPois = [...found.values()]
    .sort((a, b) => ((+a.y - cy) ** 2 + (+a.x - cx) ** 2) - ((+b.y - cy) ** 2 + (+b.x - cx) ** 2))
    .slice(0, map.getLevel() <= 2 ? 30 : 14);
  clearPois();
  for (const p of shownPois) {
    const node = document.createElement('button');
    node.type = 'button'; node.className = 'poi';
    node.textContent = p.place_name;
    node.title = (p.category_name || '').split(' > ').slice(-1)[0];
    node.onclick = ev => { ev.stopPropagation(); choosePoi(p); };
    const ov = new kakao.maps.CustomOverlay({ position: new kakao.maps.LatLng(+p.y, +p.x), content: node, yAnchor: 1.3, clickable: true, zIndex: 3 });
    ov.setMap(map); poiOverlays.push(ov);
  }
  setStatus(found.size
    ? `지도 위 이름을 누르면 그 장소가 바로 골라져요. 찾는 이름이 안 보이면 확대하거나, 지도를 눌러주세요.`
    : '이 근처에는 표시할 이름이 없어요. 간 곳을 지도에서 눌러주세요.');
}
function choosePoi(p) {
  if (!draft) return;
  stopMapPick();
  draft.lat = +p.y; draft.lng = +p.x;
  draft.place = { name: p.place_name, address: p.road_address_name || p.address_name, url: p.place_url, kakaoId: p.id };
  $('#fPlace').value = p.place_name;
  setLocText(`위치: ${p.place_name} (${draft.lat.toFixed(5)}, ${draft.lng.toFixed(5)})`);
  sg = { cur: draft, kind: null, mode: 'search', query: p.place_name, radiusIdx: 0, cats: new Set(), priority: null, showAll: false,
    pickedId: p.id, found: new Map([[p.id, { ...p, matched: false }]]) };
  renderCandidates();
  setStatus(`"${p.place_name}"을(를) 골랐어요. 다르면 다시 "지도에서 찍기"를 눌러주세요.`);
}

async function searchPlaceByName(q) {
  const cur = draft;
  if (!cur || !q) return;
  if (!settings.kakao) { setStatus('장소 검색을 쓰려면 위쪽 "설정"에 카카오 키를 넣어주세요.'); return; }
  setStatus(`"${q}" 찾는 중…`);
  try { await loadKakao(); } catch (e) { setStatus(e.message); return; }
  const S = kakao.maps.services;
  // 위치를 알면 가까운 순, 모르면 카카오가 정한 정확도 순
  const opts = cur.lat != null ? { location: new kakao.maps.LatLng(cur.lat, cur.lng), sort: S.SortBy.DISTANCE } : {};
  const data = await new Promise(res => new S.Places().keywordSearch(q, (d, st) => res(st === S.Status.OK ? d : []), opts));
  if (draft !== cur) return;
  stopMapPick();
  sg = { cur, kind: null, mode: 'search', query: q, radiusIdx: 0, cats: new Set(), priority: null, showAll: false, pickedId: null,
    found: new Map(data.map(p => [p.id, { ...p, matched: false }])) };
  renderCandidates();
}

$('#btnPlaceSearch').onclick = () => {
  const f = $('#placeSearch');
  f.hidden = !f.hidden;
  if (!f.hidden) { $('#placeQuery').value = $('#fPlace').value.trim(); $('#placeQuery').focus(); }
};
$('#placeSearch').addEventListener('submit', ev => { ev.preventDefault(); searchPlaceByName($('#placeQuery').value.trim()); });
$('#btnMapPick').onclick = () => startMapPick();

document.querySelectorAll('.slot').forEach(el => {
  const i = +el.dataset.slot;
  const open = () => { pickTarget = i; pick.value = ''; pick.click(); };
  el.addEventListener('click', e => {
    if (e.target.classList.contains('x')) { draft.photos[i] = null; draft.names[i] = null; draft.meta[i] = null; renderSlots(); return; }
    if (e.target.classList.contains('mk-main')) { swapSlots(0, i); return; }
    open();
  });
  el.addEventListener('dragstart', e => {
    if (!draft || !draft.photos[i]) { e.preventDefault(); return; }
    e.dataTransfer.setData('application/x-slot', String(i));
    e.dataTransfer.effectAllowed = 'move';
  });
  el.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag'); });
  el.addEventListener('dragleave', () => el.classList.remove('drag'));
  el.addEventListener('drop', e => {
    e.preventDefault(); el.classList.remove('drag');
    const from = e.dataTransfer.getData('application/x-slot');
    if (from !== '') { swapSlots(+from, i); return; } // 칸끼리 옮기기 = 자리 바꾸기
    const files = [...e.dataTransfer.files].filter(f => f.type.startsWith('image/'));
    files.slice(0, SLOT_COUNT - i).forEach((f, k) => setPhoto(i + k, f));
  });
});
pick.addEventListener('change', () => setPhoto(pickTarget, pick.files[0]));
// 설치한 앱(홈 화면 앱)에서는 일반 파일 선택이 갤러리로 바로 넘어가는 경우가 있어서,
// 되는 기기에서는 시스템 "파일 열기" 창(원본 그대로 받음)을 먼저 씀
$('#btnFromFiles').onclick = async () => {
  if (window.showOpenFilePicker) {
    try {
      const [h] = await window.showOpenFilePicker({
        types: [{ description: '사진', accept: { 'image/*': ['.jpg', '.jpeg', '.png', '.webp', '.heic'] } }],
        excludeAcceptAllOption: false, multiple: false,
      });
      const f = await h.getFile();
      setPhoto(0, f.type ? f : new File([f], f.name, { type: 'image/jpeg', lastModified: f.lastModified }));
      return;
    } catch (e) {
      if (e && e.name === 'AbortError') return; // 사용자가 취소
      // 지원하지 않거나 막힌 경우 아래 방식으로
    }
  }
  $('#pickFile').value = ''; $('#pickFile').click();
};
$('#pickFile').addEventListener('change', () => {
  const f = $('#pickFile').files[0];
  if (!f) return;
  if (!f.type.startsWith('image/') && !/\.(jpe?g|png|webp|heic)$/i.test(f.name)) { toast('사진 파일을 골라주세요.'); return; }
  setPhoto(0, f.type ? f : new File([f], f.name, { type: 'image/jpeg', lastModified: f.lastModified }));
});

const dbGet = id => tx('readonly', s => s.get(id));
$('#btnCancel').onclick = closeEditor;
$('#btnSave').onclick = async () => {
  if (!draft.photos[0]) { toast('대표 사진을 넣어주세요.'); return; }
  if (!$('#fDate').value) { toast('날짜를 넣어주세요.'); return; }
  const place = $('#fPlace').value.trim();
  try {
    const old = await dbGet(draft.id);
    const names = draft.photos.map((p, i) => (p ? draft.names[i] : null));
    const removed = ((old && old.photoNames) || []).filter(n => n && !names.includes(n));
    await dbPut({
      id: draft.id, created: draft.created, updated: Date.now(),
      date: $('#fDate').value, place, memo: $('#fMemo').value,
      main: draft.photos[0], extras: draft.photos.slice(1), photoNames: names,
      lat: draft.lat, lng: draft.lng, takenAt: draft.taken || null,
      placeInfo: draft.place && draft.place.name === place ? draft.place : null,
      remoteModified: old ? old.remoteModified : null,
      trash: [...((old && old.trash) || []), ...removed], dirty: true,
    });
  } catch (e) { toast('저장하지 못했어요: ' + e.message, 6000); return; }
  closeEditor(); render(); toast('저장했어요.');
  sync();
};

/* ================= 원드라이브 동기화 ================= */
// 원드라이브 앱 폴더 구조: entries/<id>.json (글), photos/<이름>.jpg (사진), settings.json (카카오 키)
const META_KEYS = ['id', 'date', 'place', 'memo', 'lat', 'lng', 'placeInfo', 'created', 'updated', 'takenAt', 'calCover'];
const newPhotoName = (id, i) => `${id}-${i}-${Math.random().toString(36).slice(2, 8)}.jpg`;
let syncing = null, lastSync = 0;

function setSyncUI(state, detail) {
  const b = $('#btnSync');
  if (!od.configured()) { b.hidden = true; return; }
  b.hidden = false;
  if (!od.signedIn()) {
    b.textContent = od.wasConnected() ? '원드라이브 다시 연결' : '원드라이브 연결';
    b.disabled = false; b.title = '원드라이브에 로그인해서 폰과 컴퓨터를 동기화해요';
    return;
  }
  b.disabled = state === 'busy';
  b.textContent = state === 'busy' ? '동기화 중…' : state === 'error' ? '동기화 실패 · 다시' : '동기화';
  b.title = detail || (lastSync ? `마지막 동기화 ${new Date(lastSync).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })}` : '');
}

async function pushEntry(e) {
  if (e.deleted) {
    if (e.remoteModified) await od.deleteFile(`entries/${e.id}.json`);
    for (const n of [...(e.photoNames || []), ...(e.trash || [])]) if (n) await od.deleteFile(`photos/${n}`);
    await dbDel(e.id);
    return;
  }
  const blobs = [e.main, ...(e.extras || [])];
  const names = [...(e.photoNames || [])];
  while (names.length < SLOT_COUNT) names.push(null);
  for (let i = 0; i < SLOT_COUNT; i++) {
    if (blobs[i] && !names[i]) {
      const n = newPhotoName(e.id, i);
      await od.putFile(`photos/${n}`, blobs[i], 'image/jpeg');
      names[i] = n;
    }
    if (!blobs[i]) names[i] = null;
  }
  const meta = { photos: names };
  for (const k of META_KEYS) meta[k] = e[k] ?? null;
  const item = await od.putFile(`entries/${e.id}.json`, JSON.stringify(meta), 'application/json');
  for (const n of e.trash || []) await od.deleteFile(`photos/${n}`).catch(() => {});
  // 그 사이에 다시 고쳐졌으면 덮어쓰지 않음
  const now = await dbGet(e.id);
  if (now && now.updated === e.updated && !now.deleted) {
    await dbPut({ ...now, photoNames: names, remoteModified: item.lastModifiedDateTime, trash: [], dirty: false });
  }
}

async function pullAll() {
  const remote = await od.listFolder('entries');
  const local = new Map((await dbAll()).map(e => [e.id, e]));
  const seen = new Set();
  let n = 0;
  for (const item of remote) {
    if (!item.name.endsWith('.json')) continue;
    const id = item.name.slice(0, -5); seen.add(id);
    const mine = local.get(id);
    if (mine && (mine.dirty || mine.remoteModified === item.lastModifiedDateTime)) continue;
    const meta = JSON.parse(await (await od.download(item)).text());
    const names = meta.photos || [];
    const oldNames = (mine && mine.photoNames) || [], oldBlobs = mine ? [mine.main, ...(mine.extras || [])] : [];
    const blobs = [];
    for (let i = 0; i < SLOT_COUNT; i++) {
      const nm = names[i];
      if (!nm) { blobs.push(null); continue; }
      const k = oldNames.indexOf(nm);
      blobs.push(k >= 0 && oldBlobs[k] ? oldBlobs[k] : await od.download(`photos/${nm}`));
    }
    const rec = { photoNames: names, main: blobs[0], extras: blobs.slice(1), remoteModified: item.lastModifiedDateTime, dirty: false, trash: [] };
    for (const k of META_KEYS) rec[k] = meta[k] ?? null;
    rec.id = id;
    await dbPut(rec);
    if (++n % 3 === 0) render();
  }
  // 다른 기기에서 지운 기록
  for (const e of local.values()) if (!e.dirty && e.remoteModified && !seen.has(e.id)) await dbDel(e.id);
}

async function sync() {
  if (!od.configured() || !od.signedIn()) { setSyncUI(); return; }
  if (syncing) return syncing;
  syncing = (async () => {
    setSyncUI('busy');
    try {
      // 아직 한 번도 올리지 않은 기록(연결 전에 쓴 것)도 함께 올림
      for (const e of await dbAll()) if (e.dirty || (!e.remoteModified && !e.deleted)) await pushEntry(e);
      await pullAll();
      await pullSettings();
      lastSync = Date.now();
      setSyncUI('idle');
    } catch (e) {
      if (e instanceof od.AuthNeeded) {
        // 쓰는 중이 아니면 화면 없이 자동으로 다시 로그인. 쓴 기록은 기기에 남아 있다가 돌아오면 올라감.
        if (editor.hidden && $('#settings').hidden && od.trySilentSignIn()) return;
        setSyncUI(); toast('원드라이브 로그인이 만료됐어요. "원드라이브 다시 연결"을 눌러주세요.', 6000);
      }
      else { setSyncUI('error', e.message); toast('동기화하지 못했어요: ' + e.message, 6000); }
    } finally { syncing = null; render(); }
  })();
  return syncing;
}

$('#btnSync').onclick = () => (od.signedIn() ? sync() : od.signIn());
window.addEventListener('focus', () => { if (Date.now() - lastSync > 30000) sync(); });
window.addEventListener('online', () => sync());

$('#btnNew').onclick = () => openEditor(null);

/* ================= 목록 ================= */
const list = $('#list');
let urls = [], current = [];
const days = ['일', '월', '화', '수', '목', '금', '토'];
async function render() {
  urls.forEach(URL.revokeObjectURL); urls = [];
  current = (await dbAll()).filter(e => !e.deleted && e.date).sort((a, b) => b.date.localeCompare(a.date) || b.created - a.created);
  const u = b => { const x = URL.createObjectURL(b); urls.push(x); return x; };
  mainUrl = new Map(current.filter(e => e.main).map(e => [e.id, u(e.main)]));
  extraUrls = new Map(current.map(e => [e.id, (e.extras || []).map(b => (b ? u(b) : null))]));
  refreshViews();
  if (!current.length) { $('#searchInfo').textContent = ''; list.innerHTML = '<p class="empty">아직 기록이 없어요. “+ 새 기록”을 눌러 첫 사진을 올려보세요.</p>'; return; }
  const q = $('#search').value.trim().toLowerCase();
  const shownList = q ? current.filter(e => matchesSearch(e, q)) : current;
  $('#searchInfo').textContent = q ? `검색 결과 ${shownList.length}개` : '';
  if (!shownList.length) { list.innerHTML = '<p class="empty">검색 결과가 없어요.</p>'; return; }
  let html = '', lastMonth = '';
  for (const e of shownList) {
    const [y, m, d] = e.date.split('-').map(Number);
    const month = `${y}년 ${m}월`;
    if (month !== lastMonth) { html += `<div class="month">${month}</div>`; lastMonth = month; }
    const dow = days[new Date(y, m - 1, d).getDay()];
    const extras = (extraUrls.get(e.id) || []).filter(Boolean).map(x => `<img src="${x}" alt="">`).join('');
    const pi = e.placeInfo;
    const addr = pi ? `<div class="addr">${esc(pi.address)}${/^https:\/\//.test(pi.url || '') ? ` · <a href="${esc(pi.url)}" target="_blank" rel="noopener">지도</a>` : ''}</div>` : '';
    html += `
      <article class="entry" data-id="${esc(e.id)}">
        ${mainUrl.has(e.id) ? `<img class="main-img" src="${mainUrl.get(e.id)}" alt="">` : '<div class="main-img"></div>'}
        <div>
          <h2>${m}월${d}일 (${dow})${e.place ? ' — ' + esc(e.place) : ''}</h2>
          ${addr}
          <div class="memo">${esc(e.memo)}</div>
        </div>
        ${extras ? `<div class="extras">${extras}</div>` : '<div></div>'}
        <div class="tools">
          <button class="small" data-act="edit">수정</button>
          <button class="small danger" data-act="del">삭제</button>
        </div>
      </article>`;
  }
  list.innerHTML = html;
}
list.addEventListener('click', async ev => {
  const art = ev.target.closest('.entry'); if (!art) return;
  const entry = current.find(x => x.id === art.dataset.id); if (!entry) return;
  const btn = ev.target.closest('button'), act = btn && btn.dataset.act;
  if (act === 'edit') openEditor(entry);
  else if (act === 'del') {
    if (!btn.dataset.armed) {
      btn.dataset.armed = '1'; btn.textContent = '정말 삭제?';
      setTimeout(() => { if (btn.isConnected) { delete btn.dataset.armed; btn.textContent = '삭제'; } }, 3000);
      return;
    }
    if (entry.remoteModified || entry.trash?.length) await dbPut({ ...entry, deleted: true, dirty: true });
    else await dbDel(entry.id);
    render(); toast('삭제했어요.'); sync();
  } else if (ev.target.tagName === 'IMG') {
    const lb = $('#lightbox'); lb.querySelector('img').src = ev.target.src; lb.hidden = false;
  }
});
$('#lightbox').onclick = () => ($('#lightbox').hidden = true);

/* ================= 검색 ================= */
function matchesSearch(e, q) {
  const [y, m, d] = e.date.split('-').map(Number);
  const hay = [e.place, e.memo, e.placeInfo && e.placeInfo.address, e.date, `${m}월${d}일`, `${m}월 ${d}일`, `${y}년 ${m}월`, `${y}년${m}월`]
    .filter(Boolean).join(' ').toLowerCase();
  return q.split(/\s+/).every(w => hay.includes(w));
}
let searchTimer;
$('#search').addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(render, 200); });

/* ================= 보기 전환 (목록 / 지도 / 캘린더) ================= */
let view = 'list', mainUrl = new Map(), extraUrls = new Map();
const slotUrl = (e, i) => (i === 0 ? mainUrl.get(e.id) : (extraUrls.get(e.id) || [])[i - 1]) || null;
function showView(v) {
  view = v;
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.view === v));
  $('#viewList').hidden = v !== 'list';
  $('#viewMap').hidden = v !== 'map';
  $('#viewCal').hidden = v !== 'cal';
  $('#viewWeek').hidden = v !== 'week';
  $('#viewYear').hidden = v !== 'year';
  try { localStorage.setItem('view', v); } catch { /* 무시 */ }
  refreshViews();
}
function refreshViews() {
  if (view === 'map') renderBigMap();
  else if (view === 'cal') renderCal();
  else if (view === 'week') renderWeek();
  else if (view === 'year') renderYear();
}
document.querySelectorAll('.tabs button').forEach(b => (b.onclick = () => showView(b.dataset.view)));

// 지도·캘린더에서 누르면 목록의 그 기록으로 이동
async function goToEntry(id) {
  showView('list');
  if ($('#search').value) { $('#search').value = ''; await render(); }
  const el = list.querySelector(`.entry[data-id="${CSS.escape(id)}"]`);
  if (!el) return;
  el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 1800);
}
async function goToDate(iso) { showView('list'); $('#search').value = iso; await render(); window.scrollTo({ top: 0, behavior: 'smooth' }); }

/* ---------- 지도 보기 ---------- */
let bigMap = null, bigOverlays = [];
async function renderBigMap() {
  const info = $('#mapInfo'), sel = $('#mapMonth');
  const withLoc = current.filter(e => e.lat != null && e.lng != null);
  const months = [...new Set(withLoc.map(e => e.date.slice(0, 7)))].sort().reverse();
  const prev = sel.value;
  sel.innerHTML = '<option value="">전체 기간</option>' + months.map(m => `<option value="${m}">${+m.slice(0, 4)}년 ${+m.slice(5)}월</option>`).join('');
  sel.value = months.includes(prev) ? prev : '';
  const items = sel.value ? withLoc.filter(e => e.date.startsWith(sel.value)) : withLoc;
  const noLoc = current.length - withLoc.length;
  if (!settings.kakao) { info.textContent = '지도를 보려면 위쪽 "설정"에 카카오 키를 넣어주세요.'; return; }
  try { await loadKakao(); } catch (e) { info.textContent = e.message; return; }
  if (view !== 'map') return;
  if (!bigMap) bigMap = new kakao.maps.Map($('#bigMap'), { center: new kakao.maps.LatLng(37.5665, 126.978), level: 8 });
  else bigMap.relayout();
  bigOverlays.forEach(o => o.setMap(null)); bigOverlays = [];
  info.textContent = `${items.length}곳` + (noLoc ? ` · 위치 정보가 없는 기록 ${noLoc}개는 지도에 안 나와요.` : '');
  if (!items.length) return;
  const bounds = new kakao.maps.LatLngBounds();
  for (const e of [...items].reverse()) {
    const pos = new kakao.maps.LatLng(e.lat, e.lng);
    bounds.extend(pos);
    const node = document.createElement('button');
    node.type = 'button'; node.className = 'pin';
    const url = mainUrl.get(e.id);
    if (url) { const img = document.createElement('img'); img.src = url; img.alt = ''; node.appendChild(img); }
    const label = document.createElement('span');
    const [, m, d] = e.date.split('-').map(Number);
    label.textContent = `${m}/${d} ${e.place || ''}`.trim();
    node.appendChild(label);
    node.onclick = () => goToEntry(e.id);
    const ov = new kakao.maps.CustomOverlay({ position: pos, content: node, yAnchor: 1, clickable: true });
    ov.setMap(bigMap); bigOverlays.push(ov);
  }
  if (items.length === 1) { bigMap.setLevel(4); bigMap.setCenter(new kakao.maps.LatLng(items[0].lat, items[0].lng)); }
  else bigMap.setBounds(bounds, 60, 40, 40, 40);
}
$('#mapMonth').onchange = () => renderBigMap();

/* ---------- 캘린더 보기 ---------- */
let calMonth = null, calAuto = false; // 'YYYY-MM'
function renderCal() {
  // 처음에는 가장 최근 기록이 있는 달을 보여줌 (기록을 불러오기 전이면 이번 달로 두었다가 바꿈)
  if (!calMonth || (calAuto && current.length)) {
    calMonth = current.length ? current[0].date.slice(0, 7) : toISO(new Date()).slice(0, 7);
    calAuto = !current.length;
  }
  const [y, m] = calMonth.split('-').map(Number);
  $('#calTitle').textContent = `${y}년 ${m}월`;
  const lead = new Date(y, m - 1, 1).getDay(), days = new Date(y, m, 0).getDate();
  const byDay = new Map();
  for (const e of current) {
    if (!e.date.startsWith(calMonth + '-')) continue;
    const d = +e.date.slice(8);
    if (!byDay.has(d)) byDay.set(d, []);
    byDay.get(d).push(e);
  }
  const todayIso = toISO(new Date());
  let html = ['일', '월', '화', '수', '목', '금', '토'].map((w, i) => `<div class="cal-dow${i === 0 ? ' sun' : i === 6 ? ' sat' : ''}">${w}</div>`).join('');
  for (let i = 0; i < lead; i++) html += '<div class="cal-cell empty"></div>';
  for (let d = 1; d <= days; d++) {
    // 기본: 그날 가장 먼저 간 곳의 대표 사진. 날짜를 꾹 눌러 직접 고른 사진(calCover)이 있으면 그걸로.
    const dayList = (byDay.get(d) || []).slice().sort((a, b) => timeKey(a).localeCompare(timeKey(b)));
    const coverE = dayList.find(x => x.calCover != null && slotUrl(x, x.calCover));
    const e = coverE || dayList[0], iso = `${calMonth}-${String(d).padStart(2, '0')}`, dow = (lead + d - 1) % 7;
    const url = e && (coverE ? slotUrl(coverE, coverE.calCover) : slotUrl(e, 0));
    html += `<div class="cal-cell${e ? ' has' : ''}${iso === todayIso ? ' today' : ''}"${e ? ` data-id="${esc(e.id)}" data-iso="${iso}" data-n="${dayList.length}" role="button" tabindex="0" title="${esc(e.place || '')}"` : ''}>`
      + (url ? `<img src="${url}" alt="">` : '')
      + `<span class="cal-d${dow === 0 ? ' sun' : dow === 6 ? ' sat' : ''}">${d}</span>`
      + (e ? `<span class="cal-place">${esc(e.place || '기록')}${dayList.length > 1 ? ` +${dayList.length - 1}` : ''}</span>` : '')
      + '</div>';
  }
  $('#calGrid').innerHTML = html;
}
function openCalCell(c) { if (!c) return; +c.dataset.n > 1 ? goToDate(c.dataset.iso) : goToEntry(c.dataset.id); }
// 날짜를 꾹 누르면(0.55초) 또는 오른쪽 클릭하면 그날 달력 사진 고르기
let lpTimer = null, lpFired = false, lpStart = null;
$('#calGrid').addEventListener('pointerdown', ev => {
  const c = ev.target.closest('.cal-cell[data-id]');
  if (!c) return;
  lpFired = false; lpStart = [ev.clientX, ev.clientY];
  clearTimeout(lpTimer);
  lpTimer = setTimeout(() => { lpFired = true; openCoverSheet(c.dataset.iso); }, 550);
});
$('#calGrid').addEventListener('pointermove', ev => {
  if (lpStart && Math.hypot(ev.clientX - lpStart[0], ev.clientY - lpStart[1]) > 10) clearTimeout(lpTimer);
});
['pointerup', 'pointercancel', 'pointerleave'].forEach(t => $('#calGrid').addEventListener(t, () => clearTimeout(lpTimer)));
$('#calGrid').addEventListener('contextmenu', ev => {
  const c = ev.target.closest('.cal-cell[data-id]');
  if (!c) return;
  ev.preventDefault(); clearTimeout(lpTimer); lpFired = true; openCoverSheet(c.dataset.iso);
});
$('#calGrid').addEventListener('click', ev => {
  if (lpFired) { lpFired = false; return; }
  openCalCell(ev.target.closest('.cal-cell[data-id]'));
});

/* ---------- 그날 달력 사진 고르기 ---------- */
let coverIso = null;
function openCoverSheet(iso) {
  const dayList = current.filter(e => e.date === iso).sort((a, b) => timeKey(a).localeCompare(timeKey(b)));
  if (!dayList.length) return;
  coverIso = iso;
  const [, m, d] = iso.split('-').map(Number);
  $('#coverTitle').textContent = `${m}월 ${d}일 달력 사진 고르기`;
  let html = '';
  for (const e of dayList) for (let i = 0; i < SLOT_COUNT; i++) {
    const url = slotUrl(e, i);
    if (!url) continue;
    const on = e.calCover === i;
    html += `<button type="button" class="cover-pick${on ? ' on' : ''}" data-id="${esc(e.id)}" data-slot="${i}">`
      + `<img src="${url}" alt=""><span>${esc(e.place || '기록')}${i ? '' : ' · 대표'}</span></button>`;
  }
  $('#coverGrid').innerHTML = html;
  $('#coverSheet').hidden = false;
}
async function setCover(iso, id, slot) {
  // 그날 기록들 중 고른 사진 하나에만 표시하고 나머지는 지움 (원드라이브로 동기화됨)
  for (const e of current.filter(x => x.date === iso)) {
    const want = e.id === id ? slot : null;
    if ((e.calCover ?? null) === want) continue;
    await dbPut({ ...e, calCover: want, dirty: true, updated: Date.now() });
  }
  $('#coverSheet').hidden = true;
  await render();
  toast(id ? '달력 사진을 바꿨어요.' : '그날 첫 기록의 대표 사진으로 되돌렸어요.');
  sync();
}
$('#coverGrid').addEventListener('click', ev => {
  const b = ev.target.closest('.cover-pick');
  if (b) setCover(coverIso, b.dataset.id, +b.dataset.slot);
});
$('#coverReset').onclick = () => setCover(coverIso, null, null);
$('#coverClose').onclick = () => ($('#coverSheet').hidden = true);
$('#coverSheet').addEventListener('click', ev => { if (ev.target.id === 'coverSheet') $('#coverSheet').hidden = true; });

/* ---------- 1년 사진첩: 대표·추가 사진 구분 없이 날짜순으로 쭉 ---------- */
let albumYear = null, albumAuto = false;
function renderYear() {
  if (!albumYear || (albumAuto && current.length)) {
    albumYear = current.length ? +current[0].date.slice(0, 4) : new Date().getFullYear();
    albumAuto = !current.length;
  }
  $('#yearTitle').textContent = `${albumYear}년`;
  const items = current.filter(e => e.date.startsWith(albumYear + '-'))
    .sort((a, b) => a.date.localeCompare(b.date) || timeKey(a).localeCompare(timeKey(b)));
  const yy = String(albumYear).slice(2);
  if (!items.length) { $('#yearList').innerHTML = `<p class="empty">${albumYear}년에는 기록이 없어요.</p>`; return; }
  let html = '', lastMonth = '';
  for (const e of items) {
    const [, m, d] = e.date.split('-');
    if (m !== lastMonth) {
      if (lastMonth) html += '</div>';
      html += `<h3 class="yr-month">${yy}.${m}</h3><div class="yr-flow">`;
      lastMonth = m;
    }
    const urls = [];
    for (let i = 0; i < SLOT_COUNT; i++) { const u = slotUrl(e, i); if (u) urls.push(u); }
    html += `<div class="yr-group"><div class="yr-date">${yy}.${m}.${d}</div><div class="yr-photos">`
      + urls.map(u => `<button type="button" class="yr-ph" data-iso="${e.date}" title="${esc(e.place || '')}"><img src="${u}" alt="" loading="lazy"></button>`).join('')
      + '</div></div>';
  }
  html += '</div>';
  $('#yearList').innerHTML = html;
}
// 사진을 누르면 캘린더의 그 날짜로
function goToCalendar(iso) {
  calMonth = iso.slice(0, 7); calAuto = false;
  showView('cal');
  const c = $(`#calGrid .cal-cell[data-iso="${iso}"]`);
  if (c) { c.scrollIntoView({ behavior: 'smooth', block: 'center' }); c.classList.add('flash'); setTimeout(() => c.classList.remove('flash'), 1800); }
}
$('#yearList').addEventListener('click', ev => { const b = ev.target.closest('.yr-ph'); if (b) goToCalendar(b.dataset.iso); });
$('#yearPrev').onclick = () => { albumYear--; albumAuto = false; renderYear(); };
$('#yearNext').onclick = () => { albumYear++; albumAuto = false; renderYear(); };
$('#calGrid').addEventListener('keydown', ev => { if (ev.key === 'Enter') openCalCell(ev.target.closest('.cal-cell[data-id]')); });
function shiftMonth(n) {
  const [y, m] = calMonth.split('-').map(Number), d = new Date(y, m - 1 + n, 1);
  calMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  calAuto = false;
  renderCal();
}
$('#calPrev').onclick = () => shiftMonth(-1);
$('#calNext').onclick = () => shiftMonth(1);
$('#calToday').onclick = () => { calMonth = toISO(new Date()).slice(0, 7); calAuto = false; renderCal(); };

/* ---------- Week 보기: 월~일, 하루 안에서는 찍은 시간 순 ---------- */
const DOW_EN = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const isoToDate = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d); };
const addDays = (iso, n) => { const d = isoToDate(iso); d.setDate(d.getDate() + n); return toISO(d); };
const mondayOf = iso => addDays(iso, -((isoToDate(iso).getDay() + 6) % 7));
const dotted = iso => iso.replace(/-/g, '.');
// 찍은 시각이 있으면 그걸로, 없으면(예전 기록) 기록을 쓴 시각으로 정렬
const timeKey = e => e.takenAt || (e.created ? `${e.date}T${new Date(e.created).toTimeString().slice(0, 8)}` : `${e.date}T99`);
let weekStart = null, weekAuto = false;

function renderWeek() {
  if (!weekStart || (weekAuto && current.length)) {
    weekStart = mondayOf(current.length ? current[0].date : toISO(new Date()));
    weekAuto = !current.length;
  }
  $('#weekTitle').textContent = `${dotted(weekStart)} ~ ${dotted(addDays(weekStart, 6))}`;
  const today = toISO(new Date());
  let html = '';
  for (let i = 0; i < 7; i++) {
    const iso = addDays(weekStart, i), [, m, d] = iso.split('-').map(Number);
    const items = current.filter(e => e.date === iso).sort((a, b) => timeKey(a).localeCompare(timeKey(b)));
    html += `<section class="wk-day${items.length ? '' : ' wk-empty'}${iso === today ? ' today' : ''}">
      <h3><span class="wk-dow${i === 5 ? ' sat' : i === 6 ? ' sun' : ''}">${DOW_EN[i]}</span><span class="wk-date">${m}.${d}</span>${items.length ? `<span class="wk-count">${items.length}곳</span>` : ''}</h3>`;
    if (items.length) {
      html += '<div class="wk-items">' + items.map((e, k) => {
        const url = mainUrl.get(e.id);
        const time = e.takenAt ? e.takenAt.slice(11, 16) : '';
        return (k ? '<div class="wk-step" aria-hidden="true">↓</div>' : '')
          + `<button type="button" class="wk-item" data-id="${esc(e.id)}">`
          + (url ? `<img src="${url}" alt="">` : '<div class="ph"></div>')
          + `<div class="wk-text"><div class="wk-place">${time ? `<span class="wk-time">${time}</span>` : ''}${esc(e.place || '장소 없음')}</div>`
          + (e.memo ? `<div class="wk-memo">${esc(e.memo)}</div>` : '')
          + '</div></button>';
      }).join('') + '</div>';
    }
    html += '</section>';
  }
  $('#weekList').innerHTML = html;
}
$('#weekList').addEventListener('click', ev => { const b = ev.target.closest('.wk-item'); if (b) goToEntry(b.dataset.id); });
$('#weekPrev').onclick = () => { weekStart = addDays(weekStart, -7); weekAuto = false; renderWeek(); };
$('#weekNext').onclick = () => { weekStart = addDays(weekStart, 7); weekAuto = false; renderWeek(); };
$('#weekNow').onclick = () => { weekStart = mondayOf(toISO(new Date())); weekAuto = false; renderWeek(); };
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('#lightbox').hidden = true; });

/* ================= 백업 ================= */
const blobToData = b => b ? new Promise(r => { const f = new FileReader(); f.onload = () => r(f.result); f.readAsDataURL(b); }) : null;
const dataToBlob = s => s ? fetch(s).then(r => r.blob()) : null;
$('#btnExport').onclick = async () => {
  const out = [];
  for (const e of await dbAll()) {
    if (e.deleted) continue;
    const { photoNames, remoteModified, trash, dirty, ...rest } = e;
    out.push({ ...rest, main: await blobToData(e.main), extras: await Promise.all((e.extras || []).map(blobToData)) });
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify({ version: 2, entries: out })], { type: 'application/json' }));
  a.download = `사진일기-백업-${toISO(new Date())}.json`;
  a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
};
$('#btnImport').onclick = () => { $('#importFile').value = ''; $('#importFile').click(); };
$('#importFile').onchange = async e => {
  const f = e.target.files[0]; if (!f) return;
  try {
    const data = JSON.parse(await f.text());
    for (const x of data.entries || []) {
      const extras = await Promise.all((x.extras || []).slice(0, EXTRA_COUNT).map(dataToBlob));
      while (extras.length < EXTRA_COUNT) extras.push(null);
      await dbPut({ ...x, main: await dataToBlob(x.main), extras, photoNames: [], remoteModified: null, trash: [], dirty: true, updated: Date.now() });
    }
    render(); toast(`${(data.entries || []).length}개 기록을 불러왔어요.`);
    sync();
  } catch { toast('백업 파일을 읽지 못했어요.'); }
};

/* ================= 시작 ================= */
updateKeyNote();
setSyncUI();
render().catch(() => { list.innerHTML = '<p class="empty">이 브라우저에서는 저장 공간을 쓸 수 없어요.</p>'; });
try { const v = localStorage.getItem('view'); if (['map', 'cal', 'week', 'year'].includes(v)) showView(v); } catch { /* 무시 */ }
od.handleRedirect()
  .then(justSignedIn => {
    if (justSignedIn && !sessionStorage.getItem('silentTriedAt')) toast('원드라이브에 연결했어요.');
    if (!od.signedIn() && od.trySilentSignIn()) return; // 만료됐으면 앱을 열 때 자동으로 다시 로그인
    sync();
  })
  .catch(e => { setSyncUI(); toast('원드라이브 로그인 실패: ' + e.message, 6000); });
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
