// POST /api/proposal — 입력값 검사 → R2 에 상태 기록 → GitHub Actions(proposal.yml) 실행
import { json, requireAccess } from '../../_shared.js';

const DATE = /^\d{4}-\d{2}-\d{2}$/;
const TEXT_MAX = 1500;

function clean(body) {
  const s = v => (v == null ? '' : String(v)).trim();
  const out = {
    corp: s(body.corp), issue: s(body.issue), demand: s(body.demand), base: s(body.base), max: s(body.max),
    split: s(body.split).replace(/\s/g, ''), priority: s(body.priority), band: s(body.band).slice(0, 80), fee_bp: s(body.fee_bp).slice(0, 10),
    issues_note: s(body.issues_note).slice(0, TEXT_MAX), credit_note: s(body.credit_note).slice(0, TEXT_MAX),
  };
  if (!/^[0-9A-Za-z]{1,20}$/.test(out.corp)) return [null, '회사 코드가 올바르지 않습니다.'];
  if (!DATE.test(out.issue)) return [null, '발행일을 입력해 주세요.'];
  if (out.demand && !DATE.test(out.demand)) return [null, '수요예측일 형식이 올바르지 않습니다.'];
  if (!/^\d{1,2}:\d{1,6}(,\d{1,2}:\d{1,6}){0,4}$/.test(out.split)) return [null, '만기별 금액 형식이 올바르지 않습니다 (예: 2:600,3:1500).'];
  for (const k of ['base', 'max', 'priority']) if (out[k] && !/^\d{1,6}$/.test(out[k])) return [null, `${k} 값이 올바르지 않습니다.`];
  if (out.fee_bp && !/^\d{1,3}(\.\d{1,2})?$/.test(out.fee_bp)) return [null, '인수수수료는 숫자(bp)로 입력해 주세요.'];
  const sum = out.split.split(',').reduce((a, p) => a + Number(p.split(':')[1]), 0);
  if (!out.base) out.base = String(sum);
  if (Number(out.base) !== sum) return [null, '만기별 금액 합계가 발행금액과 다릅니다.'];
  return [out, null];
}

export async function onRequestPost({ request, env }) {
  const denied = requireAccess(request, env);
  if (denied) return denied;
  if (!env.GH_TOKEN || !env.GH_REPO || !env.PROPOSALS) return json({ error: '서버 설정이 끝나지 않았습니다 (GH_TOKEN·GH_REPO·R2 바인딩).' }, 500);
  let body;
  try { body = await request.json(); } catch { return json({ error: '잘못된 요청입니다.' }, 400); }
  const [inp, err] = clean(body);
  if (err) return json({ error: err }, 400);

  const id = crypto.randomUUID();
  const user = request.headers.get('cf-access-authenticated-user-email') || '';
  await env.PROPOSALS.put(`proposals/${id}/status.json`, JSON.stringify({ state: 'queued', corp: inp.corp, user, at: new Date().toISOString() }),
    { httpMetadata: { contentType: 'application/json' } });

  const r = await fetch(`https://api.github.com/repos/${env.GH_REPO}/actions/workflows/proposal.yml/dispatches`, {
    method: 'POST',
    headers: { authorization: `Bearer ${env.GH_TOKEN}`, accept: 'application/vnd.github+json', 'user-agent': 'ecm-board-proposal', 'x-github-api-version': '2022-11-28' },
    body: JSON.stringify({ ref: env.GH_REF || 'main', inputs: { request_id: id, payload: JSON.stringify(inp) } }),
  });
  if (!r.ok) {
    const msg = `GitHub 실행 요청 실패 (${r.status})`;
    await env.PROPOSALS.put(`proposals/${id}/status.json`, JSON.stringify({ state: 'error', error: msg }));
    return json({ error: msg }, 502);
  }
  return json({ id });
}
