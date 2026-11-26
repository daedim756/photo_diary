// 원드라이브(개인 계정) 연결: 로그인(OAuth PKCE) + 앱 폴더 파일 읽기/쓰기
import { MS_CLIENT_ID } from './config.js';

const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
// Files.Read: 카메라 백업 사진에서 위치를 읽기 위해 (읽기 전용)
const SCOPE = 'Files.ReadWrite.AppFolder Files.Read offline_access';
const DRIVE = 'https://graph.microsoft.com/v1.0/me/drive';
const ROOT = 'https://graph.microsoft.com/v1.0/me/drive/special/approot';
const TOKEN_KEY = 'msToken';

export const configured = () => !!MS_CLIENT_ID;
const redirectUri = () => new URL('./', location.href).href;

const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randomStr = n => b64url(crypto.getRandomValues(new Uint8Array(n)));

function loadToken() { try { return JSON.parse(localStorage.getItem(TOKEN_KEY) || 'null'); } catch { return null; } }
function saveToken(t) {
  localStorage.setItem(TOKEN_KEY, JSON.stringify({
    access: t.access_token, refresh: t.refresh_token, expires: Date.now() + (t.expires_in - 120) * 1000,
  }));
  localStorage.setItem('msConnected', '1');
}
export const signedIn = () => !!loadToken();
// 토큰은 만료돼도 "예전에 연결했던 기기"인지는 기억해서 자동 재로그인에 씀
export const wasConnected = () => localStorage.getItem('msConnected') === '1';
export function signOut() { localStorage.removeItem(TOKEN_KEY); }

export async function signIn({ silent = false } = {}) {
  const verifier = randomStr(48), state = randomStr(16);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  sessionStorage.setItem('pkce', JSON.stringify({ verifier, state, silent }));
  const q = new URLSearchParams({
    client_id: MS_CLIENT_ID, response_type: 'code', redirect_uri: redirectUri(), scope: SCOPE,
    code_challenge: challenge, code_challenge_method: 'S256', state,
    prompt: silent ? 'none' : 'select_account', // none: Microsoft에 이미 로그인돼 있으면 화면 없이 바로 돌아옴
  });
  location.href = `${AUTH}/authorize?${q}`;
}

// 로그인이 만료됐을 때 화면 없이 다시 로그인을 시도. 10분에 한 번만 (반복 방지).
export function trySilentSignIn() {
  if (!configured() || signedIn() || !wasConnected()) return false;
  const last = +(sessionStorage.getItem('silentTriedAt') || 0);
  if (Date.now() - last < 10 * 60 * 1000) return false;
  sessionStorage.setItem('silentTriedAt', String(Date.now()));
  signIn({ silent: true });
  return true;
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: MS_CLIENT_ID, scope: SCOPE, ...params }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.error || `로그인 실패 (${res.status})`);
  saveToken(data);
}

// 로그인 후 돌아왔을 때 주소의 ?code= 를 토큰으로 바꿈. 처리했으면 true.
export async function handleRedirect() {
  const p = new URLSearchParams(location.search);
  if (!p.has('code') && !p.has('error')) return false;
  history.replaceState(null, '', location.pathname);
  const saved = JSON.parse(sessionStorage.getItem('pkce') || 'null');
  sessionStorage.removeItem('pkce');
  if (p.has('error')) {
    if (saved && saved.silent) return false; // 자동 재로그인 실패: 조용히 넘어가고 버튼으로 안내
    throw new Error(p.get('error_description') || p.get('error'));
  }
  if (!saved || saved.state !== p.get('state')) throw new Error('로그인 확인값이 맞지 않아요. 다시 로그인해주세요.');
  await tokenRequest({ grant_type: 'authorization_code', code: p.get('code'), redirect_uri: redirectUri(), code_verifier: saved.verifier });
  return true;
}

export class AuthNeeded extends Error {}
export class NeedsConsent extends Error {}
async function accessToken() {
  const t = loadToken();
  if (!t) throw new AuthNeeded('로그인이 필요해요');
  if (Date.now() < t.expires) return t.access;
  try { await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh }); }
  catch { signOut(); throw new AuthNeeded('로그인이 만료됐어요'); }
  return loadToken().access;
}

async function graph(url, opts = {}) {
  const res = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: `Bearer ${await accessToken()}` } });
  if (res.status === 401) { signOut(); throw new AuthNeeded('로그인이 만료됐어요'); }
  return res;
}
const enc = p => p.split('/').map(encodeURIComponent).join('/');

export async function putFile(path, body, type) {
  const res = await graph(`${ROOT}:/${enc(path)}:/content`, { method: 'PUT', headers: { 'Content-Type': type }, body });
  if (!res.ok) throw new Error(`원드라이브 저장 실패 (${res.status})`);
  return res.json();
}
export async function deleteFile(path) {
  const res = await graph(`${ROOT}:/${enc(path)}`, { method: 'DELETE' });
  if (!res.ok && res.status !== 404) throw new Error(`원드라이브 삭제 실패 (${res.status})`);
}
export async function listFolder(folder) {
  let url = `${ROOT}:/${enc(folder)}:/children?$top=200`, out = [];
  while (url) {
    const res = await graph(url);
    if (res.status === 404) return out;
    if (!res.ok) throw new Error(`원드라이브 목록 실패 (${res.status})`);
    const data = await res.json();
    out = out.concat(data.value || []);
    url = data['@odata.nextLink'];
  }
  return out;
}
/* ---------- 카메라 백업에서 같은 사진 찾기 ---------- */
// 폴더 하나를 최근 것부터 훑으며 촬영 시각이 같은 사진을 찾음. 하위 폴더 목록도 함께 돌려줌.
async function scanFolder(id, target, isMatch) {
  let url = `${DRIVE}/items/${id}/children?$top=200&$orderby=lastModifiedDateTime desc`, ordered = true;
  const folders = [];
  for (let page = 0; url && page < 20; page++) {
    let res = await graph(url);
    if (res.status === 400 && ordered) { ordered = false; url = `${DRIVE}/items/${id}/children?$top=200`; res = await graph(url); }
    if (res.status === 403) { signOut(); throw new NeedsConsent('원드라이브 사진을 읽을 권한이 필요해요'); }
    if (!res.ok) throw new Error(`원드라이브 사진 목록 실패 (${res.status})`);
    const data = await res.json(), items = data.value || [];
    for (const it of items) {
      if (it.folder) folders.push(it);
      else if (isMatch(it)) return { item: it, folders };
    }
    // 최근 수정 순이면, 촬영보다 한참 전에 올라온 파일부터는 볼 필요 없음
    const last = items[items.length - 1];
    if (ordered && last && Date.parse(last.lastModifiedDateTime) < target - 3 * 86400000) break;
    url = data['@odata.nextLink'];
  }
  return { item: null, folders };
}

// taken: 'YYYY-MM-DDTHH:MM:SS'(사진에 적힌 촬영 시각), offset: '+09:00' 또는 null
// 찾으면 {lat, lng} 또는 {blob}(원본 파일), 못 찾으면 null
export async function findPhotoLocation(taken, offset) {
  const target = offset ? Date.parse(taken + offset) : new Date(taken).getTime();
  const isMatch = it => {
    const t = it.photo && it.photo.takenDateTime;
    if (!t) return false;
    // 원드라이브는 시간대 없이 적힌 시각을 그대로 UTC처럼 기록하기도 해서 두 방식 모두 비교
    return Math.abs(Date.parse(t) - target) <= 2000 || t.slice(0, 19) === taken;
  };
  const res = await graph(`${DRIVE}/special/cameraroll`);
  if (res.status === 404) return null;
  if (res.status === 403) { signOut(); throw new NeedsConsent('원드라이브 사진을 읽을 권한이 필요해요'); }
  if (!res.ok) throw new Error(`원드라이브 카메라 폴더 확인 실패 (${res.status})`);
  const root = await res.json();
  let { item, folders } = await scanFolder(root.id, target, isMatch);
  // 연도/월 폴더로 정리해 두는 경우 (예: 2026 / 11)
  const [y, mo] = taken.split('-');
  for (const wanted of [[y], [mo, String(+mo), `${y}-${mo}`, `${y}${mo}`, `${y}.${mo}`]]) {
    if (item) break;
    const f = folders.find(x => wanted.includes(x.name));
    if (!f) break;
    ({ item, folders } = await scanFolder(f.id, target, isMatch));
  }
  if (!item) return null;
  const loc = item.location;
  if (loc && loc.latitude != null && loc.longitude != null) return { lat: loc.latitude, lng: loc.longitude };
  return { blob: await download(item) }; // 위치가 따로 없으면 원본 파일에서 읽음
}

// 목록에서 받은 항목(또는 경로)의 내용을 받아옴
export async function download(itemOrPath) {
  let item = itemOrPath;
  if (typeof item === 'string') {
    const res = await graph(`${ROOT}:/${enc(item)}`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`원드라이브 읽기 실패 (${res.status})`);
    item = await res.json();
  }
  const res = await fetch(item['@microsoft.graph.downloadUrl']);
  if (!res.ok) throw new Error(`원드라이브 받기 실패 (${res.status})`);
  return res.blob();
}
