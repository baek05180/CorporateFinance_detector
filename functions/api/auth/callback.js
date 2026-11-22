// GET /api/auth/callback — GitHub 가 돌려보내는 곳. 아이디가 허용 목록에 있으면 12시간 로그인 쿠키 발급.
// GitHub 사용자 토큰은 아이디 확인에만 쓰고 저장하지 않는다.
import { allowed, cookie, makeSession, setCookie, COOKIE, STATE_COOKIE } from '../../_shared.js';

const page = (msg, status = 403) => new Response(
  `<!doctype html><meta charset="utf-8"><title>로그인</title><body style="font:14px sans-serif;padding:24px">${msg}</body>`,
  { status, headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'set-cookie': setCookie(STATE_COOKIE, '', 0) } });

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const saved = cookie(request, STATE_COOKIE) || '';
  const [state, ret] = [saved.slice(0, saved.indexOf('|')), saved.slice(saved.indexOf('|') + 1)];
  if (!state || state !== url.searchParams.get('state')) return page('로그인 요청이 만료되었습니다. 창을 닫고 다시 시도해 주세요.', 400);
  const tok = await fetch('https://github.com/login/oauth/access_token', {
    method: 'POST',
    headers: { accept: 'application/json', 'content-type': 'application/json', 'user-agent': 'ecm-board-proposal' },
    body: JSON.stringify({ client_id: env.GITHUB_CLIENT_ID, client_secret: env.GITHUB_CLIENT_SECRET, code: url.searchParams.get('code'), redirect_uri: `${url.origin}/api/auth/callback` }),
  }).then(r => r.json()).catch(() => ({}));
  if (!tok.access_token) return page('GitHub 로그인에 실패했습니다.', 400);
  const user = await fetch('https://api.github.com/user', {
    headers: { authorization: `Bearer ${tok.access_token}`, accept: 'application/vnd.github+json', 'user-agent': 'ecm-board-proposal' },
  }).then(r => r.json()).catch(() => ({}));
  if (!allowed(env, user.login)) return page('사용 권한이 없는 GitHub 계정입니다.');
  const h = new Headers({ location: /^\/proposal\//.test(ret) ? ret : '/proposal/', 'cache-control': 'no-store' });
  h.append('set-cookie', setCookie(COOKIE, await makeSession(env, user.login), 12 * 3600));
  h.append('set-cookie', setCookie(STATE_COOKIE, '', 0));
  return new Response(null, { status: 302, headers: h });
}
