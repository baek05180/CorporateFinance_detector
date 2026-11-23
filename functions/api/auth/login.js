// GET /api/auth/login?return=/proposal/?corp=... → GitHub 로그인 화면으로
import { json, setCookie, STATE_COOKIE } from '../../_shared.js';

export async function onRequestGet({ request, env }) {
  if (!env.GITHUB_CLIENT_ID) return json({ error: '서버 설정이 끝나지 않았습니다 (GITHUB_CLIENT_ID).' }, 500);
  const url = new URL(request.url);
  let ret = url.searchParams.get('return') || '/proposal/';
  if (!/^\/proposal\/(\?corp=[0-9A-Za-z]{1,20})?$/.test(ret)) ret = '/proposal/';
  const state = crypto.randomUUID();
  const go = new URL('https://github.com/login/oauth/authorize');
  go.searchParams.set('client_id', env.GITHUB_CLIENT_ID);
  go.searchParams.set('redirect_uri', `${url.origin}/api/auth/callback`);
  go.searchParams.set('state', state);
  go.searchParams.set('scope', '');  // 공개 프로필(아이디)만 확인
  go.searchParams.set('allow_signup', 'false');
  return new Response(null, { status: 302, headers: { location: go.toString(), 'set-cookie': setCookie(STATE_COOKIE, `${state}|${ret}`, 600), 'cache-control': 'no-store' } });
}
