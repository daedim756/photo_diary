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
  if (ifd0[0x8769]) { const ex = ifd(ifd0[0x8769]); dt = ex[0x9003] || ex[0x9004]; }
  dt = dt || ifd0[0x0132];
  const m = dt && dt.match(/^(\d{4}):(\d{2}):(\d{2})/);
  if (m && m[1] !== '0000') out.date = `${m[1]}-${m[2]}-${m[3]}`;
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
  for (const c of onlyCats || [...s.cats]) add(await placesCall('categorySearch', c, lat, lng, r), false);
}
const sortedCandidates = s => [...s.found.values()].sort((a, b) =>
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
    photos: photos.slice(0, SLOT_COUNT), sizes: [], names: names.slice(0, SLOT_COUNT),
    lat: entry?.lat ?? null, lng: entry?.lng ?? null, place: entry?.placeInfo || null,
  };
  $('#fDate').value = entry ? entry.date : '';
  $('#fPlace').value = entry ? entry.place || '' : '';
  $('#fMemo').value = entry ? entry.memo || '' : '';
  setHint(entry ? '' : '대표 사진을 올리면 촬영 날짜가 자동으로 들어갑니다.');
  clearSuggestions();
  setLocText();
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
      el.innerHTML = `<img src="${el._url}" alt=""><button class="x" title="빼기" aria-label="사진 빼기">✕</button>`
        + (draft.sizes[i] ? `<span class="size">${draft.sizes[i]}</span>` : '');
      el.classList.add('filled');
    } else { el.innerHTML = slotLabels[i]; el.classList.remove('filled'); }
  });
}

async function setPhoto(i, file) {
  if (!draft || !file || !file.type.startsWith('image/')) return;
  let exif = {};
  if (i === 0) {
    exif = await readExif(file);
    if (exif.date) { $('#fDate').value = exif.date; setHint('사진의 촬영 날짜를 읽었습니다.', true); }
    else { $('#fDate').value = toISO(new Date(file.lastModified)); setHint('사진에 촬영 날짜 정보가 없어 파일 날짜를 넣었습니다. 필요하면 고쳐주세요.'); }
  }
  try { draft.photos[i] = await shrink(file); }
  catch { toast('이 사진 형식은 열 수 없어요. 아이폰 HEIC 사진이면 JPG로 바꿔서 올려주세요.'); return; }
  draft.sizes[i] = `${fmtSize(file.size)} → ${fmtSize(draft.photos[i].size)}`;
  draft.names[i] = null; // 새 사진: 원드라이브에 새 이름으로 올림
  renderSlots();
  if (i === 0) {
    if (exif.lat != null) { draft.lat = exif.lat; draft.lng = exif.lng; setLocText(); suggest(); }
    else {
      draft.lat = draft.lng = null; clearSuggestions();
      setLocText('위치: 이 사진에는 위치 정보가 없어요. 폰에서 올렸다면 "파일에서 대표 사진 고르기"로 다시 골라보거나, 지금 그 장소라면 "현재 위치 사용"을 눌러주세요.');
    }
  }
}

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
async function suggest() {
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
  if (kind && kind.category !== 'OTHER') { s.cats.add(kind.category); s.priority = kind.category; }
  setStatus(kind ? `사진: ${kind.subject} → ${kind.place_type} 주변에서 찾는 중…` : '주변 장소를 찾는 중…');
  try { await runSearches(s); }
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
  setStatus(all.length
    ? `${s.kind ? `사진: ${s.kind.subject} → ${s.kind.place_type}. ` : ''}반경 ${r}m 안에서 ${all.length}곳을 찾았어요.`
    : `반경 ${r}m 안에서 찾은 곳이 없어요. 범위를 넓히거나 장소를 직접 적어주세요.`);
  $('#chips').innerHTML = `<span class="chips-label">앞으로 올리기</span>`
    + CATS.map(([c, n]) => `<button type="button" class="chip${s.priority === c ? ' on' : ''}" data-cat="${c}">${n}</button>`).join('')
    + (s.radiusIdx < RADII.length - 1 ? `<button type="button" class="chip" data-wide="1">범위 넓히기 → ${RADII[s.radiusIdx + 1]}m</button>` : '');
  const ol = $('#candidates');
  ol.innerHTML = shown.map((p, k) => `
    <li><button type="button" data-k="${k}"${s.pickedId === p.id ? ' class="picked"' : ''}>
      <span class="name">${esc(p.place_name)}${p.matched && s.kind ? `<span class="tag">${esc(s.kind.place_type)}</span>` : ''}</span>
      <span class="dist">${esc(p.distance)}m</span>
      <span class="meta">${esc((p.category_name || '').split(' > ').slice(-2).join(' · '))} · ${esc(p.road_address_name || p.address_name)}</span>
    </button></li>`).join('')
    + (all.length > shown.length ? `<li><button type="button" class="more" data-more="1">${all.length - shown.length}곳 더 보기</button></li>` : '');
  ol.querySelectorAll('button[data-k]').forEach(b => b.onclick = () => {
    const p = shown[+b.dataset.k];
    $('#fPlace').value = p.place_name;
    draft.place = { name: p.place_name, address: p.road_address_name || p.address_name, url: p.place_url, kakaoId: p.id };
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
  el.hidden = false;
  const center = new kakao.maps.LatLng(draft.lat, draft.lng);
  if (!map) map = new kakao.maps.Map(el, { center, level: 3 });
  else { map.relayout(); map.setCenter(center); }
  markers.forEach(m => m.setMap(null)); markers = [];
  markers.push(new kakao.maps.Circle({ center, radius: 12, strokeWeight: 2, strokeColor: '#2e6bd6', fillColor: '#2e6bd6', fillOpacity: 0.6, map }));
  list.forEach(p => markers.push(new kakao.maps.Marker({ position: new kakao.maps.LatLng(+p.y, +p.x), title: p.place_name, map })));
}

document.querySelectorAll('.slot').forEach(el => {
  const i = +el.dataset.slot;
  const open = () => { pickTarget = i; pick.value = ''; pick.click(); };
  el.addEventListener('click', e => {
    if (e.target.classList.contains('x')) { draft.photos[i] = null; draft.names[i] = null; renderSlots(); return; }
    open();
  });
  el.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('drag'); });
  el.addEventListener('dragleave', () => el.classList.remove('drag'));
  el.addEventListener('drop', e => {
    e.preventDefault(); el.classList.remove('drag');
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
      lat: draft.lat, lng: draft.lng,
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
const META_KEYS = ['id', 'date', 'place', 'memo', 'lat', 'lng', 'placeInfo', 'created', 'updated'];
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
    const extras = (e.extras || []).filter(Boolean).map(b => `<img src="${u(b)}" alt="">`).join('');
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
let view = 'list', mainUrl = new Map();
function showView(v) {
  view = v;
  document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.view === v));
  $('#viewList').hidden = v !== 'list';
  $('#viewMap').hidden = v !== 'map';
  $('#viewCal').hidden = v !== 'cal';
  try { localStorage.setItem('view', v); } catch { /* 무시 */ }
  refreshViews();
}
function refreshViews() {
  if (view === 'map') renderBigMap();
  else if (view === 'cal') renderCal();
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
    const dayList = (byDay.get(d) || []).slice().reverse(); // 그날 먼저 쓴 기록이 대표
    const e = dayList[0], iso = `${calMonth}-${String(d).padStart(2, '0')}`, dow = (lead + d - 1) % 7;
    const url = e && mainUrl.get(e.id);
    html += `<div class="cal-cell${e ? ' has' : ''}${iso === todayIso ? ' today' : ''}"${e ? ` data-id="${esc(e.id)}" data-iso="${iso}" data-n="${dayList.length}" role="button" tabindex="0" title="${esc(e.place || '')}"` : ''}>`
      + (url ? `<img src="${url}" alt="">` : '')
      + `<span class="cal-d${dow === 0 ? ' sun' : dow === 6 ? ' sat' : ''}">${d}</span>`
      + (e ? `<span class="cal-place">${esc(e.place || '기록')}${dayList.length > 1 ? ` +${dayList.length - 1}` : ''}</span>` : '')
      + '</div>';
  }
  $('#calGrid').innerHTML = html;
}
function openCalCell(c) { if (!c) return; +c.dataset.n > 1 ? goToDate(c.dataset.iso) : goToEntry(c.dataset.id); }
$('#calGrid').addEventListener('click', ev => openCalCell(ev.target.closest('.cal-cell[data-id]')));
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
try { const v = localStorage.getItem('view'); if (v === 'map' || v === 'cal') showView(v); } catch { /* 무시 */ }
od.handleRedirect()
  .then(justSignedIn => {
    if (justSignedIn && !sessionStorage.getItem('silentTriedAt')) toast('원드라이브에 연결했어요.');
    if (!od.signedIn() && od.trySilentSignIn()) return; // 만료됐으면 앱을 열 때 자동으로 다시 로그인
    sync();
  })
  .catch(e => { setSyncUI(); toast('원드라이브 로그인 실패: ' + e.message, 6000); });
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
