// 원드라이브(개인 계정) 연결: 로그인(OAuth PKCE) + 앱 폴더 파일 읽기/쓰기
import { MS_CLIENT_ID } from './config.js';

const AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const SCOPE = 'Files.ReadWrite.AppFolder offline_access';
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
