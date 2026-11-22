// 공모사채 제안서 API 공통 함수 — 결제 없는 구성
//   로그인: GitHub OAuth (ALLOWED_GITHUB_USERS 에 있는 아이디만)
//   생성:   비공개 저장소(VAULT_REPO)의 GitHub Actions — 입력값·실행 기록·결과물 모두 비공개
//
// Cloudflare Pages → Settings → Variables and Secrets
//   GITHUB_CLIENT_ID          GitHub OAuth App 의 Client ID
//   GITHUB_CLIENT_SECRET (S)  GitHub OAuth App 의 Client secret
//   SESSION_SECRET (S)        로그인 쿠키 서명용 임의 문자열 (32자 이상)
//   ALLOWED_GITHUB_USERS      쓸 수 있는 GitHub 아이디 (쉼표 구분)
//   GH_TOKEN (S)              fine-grained 토큰 — VAULT_REPO 하나, Actions: Read and write
//   VAULT_REPO                예: baek05180/proposal-vault
//   VAULT_REF                 기본 main
export const json = (obj, status = 200, headers = {}) =>
  new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...headers } });

const enc = new TextEncoder();
const b64u = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

async function hmac(secret, data) {
  const key = await crypto.subtle.importKey('raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return b64u(await crypto.subtle.sign('HMAC', key, enc.encode(data)));
}

function safeEq(a, b) {
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

export const COOKIE = '__Host-pv_session';
export const STATE_COOKIE = '__Host-pv_state';
const SESSION_HOURS = 12;

export function cookie(request, name) {
  const m = (request.headers.get('cookie') || '').match(new RegExp(`(?:^|;\\s*)${name.replace(/[-]/g, '\\-')}=([^;]+)`));
  return m ? decodeURIComponent(m[1]) : null;
}

export const setCookie = (name, value, maxAge) =>
  `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${maxAge}`;

export function allowed(env, login) {
  const list = String(env.ALLOWED_GITHUB_USERS || '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
  return !!login && list.includes(String(login).toLowerCase());
}

export async function makeSession(env, login) {
  const exp = Math.floor(Date.now() / 1000) + SESSION_HOURS * 3600;
  const body = `${login}|${exp}`;
  return `${body}|${await hmac(env.SESSION_SECRET, body)}`;
}

// 로그인 확인. 성공하면 { login }, 실패하면 Response.
export async function authorize(request, env, { write = false } = {}) {
  if (!env.SESSION_SECRET || env.SESSION_SECRET.length < 32) return json({ error: '서버 설정이 끝나지 않았습니다 (SESSION_SECRET).' }, 500);
  // 다른 사이트에서 보낸 요청(CSRF) 차단
  if (write) {
    const origin = request.headers.get('origin');
    if (!origin || origin !== new URL(request.url).origin) return json({ error: '잘못된 요청입니다.' }, 403);
  }
  const v = cookie(request, COOKIE);
  if (!v) return json({ error: '로그인이 필요합니다.', login: true }, 401);
  const [login, exp, sig] = v.split('|');
  if (!login || !exp || !sig || !safeEq(sig, await hmac(env.SESSION_SECRET, `${login}|${exp}`)) || Number(exp) < Date.now() / 1000)
    return json({ error: '로그인이 만료되었습니다.', login: true }, 401);
  if (!allowed(env, login)) return json({ error: '사용 권한이 없는 계정입니다.' }, 403);
  return { login };
}

export function gh(env, path, init = {}) {
  return fetch(`https://api.github.com${path}`, {
    ...init,
    headers: { authorization: `Bearer ${env.GH_TOKEN}`, accept: 'application/vnd.github+json', 'user-agent': 'ecm-board-proposal', 'x-github-api-version': '2022-11-28', ...(init.headers || {}) },
  });
}
