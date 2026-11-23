// GET /api/auth/me → { login } 또는 401 ;  POST /api/auth/me → 로그아웃
import { authorize, json, setCookie, COOKIE } from '../../_shared.js';

export async function onRequestGet({ request, env }) {
  const who = await authorize(request, env);
  return who instanceof Response ? who : json({ login: who.login });
}

export async function onRequestPost({ request, env }) {
  const who = await authorize(request, env, { write: true });
  if (who instanceof Response && who.status !== 401) return who;
  return json({ ok: true }, 200, { 'set-cookie': setCookie(COOKIE, '', 0) });
}
