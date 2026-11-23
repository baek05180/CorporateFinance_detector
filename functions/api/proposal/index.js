// POST /api/proposal — 입력값 검사 → 비공개 저장소(VAULT_REPO)의 proposal.yml 실행
// 입력값은 비공개 저장소의 실행 입력으로만 전달된다 (저장소 관리자만 볼 수 있음).
import { authorize, gh, json } from '../../_shared.js';

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
  const who = await authorize(request, env, { write: true });
  if (who instanceof Response) return who;
  if (!env.GH_TOKEN || !env.VAULT_REPO) return json({ error: '서버 설정이 끝나지 않았습니다 (GH_TOKEN·VAULT_REPO).' }, 500);
  let body;
  try { body = await request.json(); } catch { return json({ error: '잘못된 요청입니다.' }, 400); }
  const [inp, err] = clean(body);
  if (err) return json({ error: err }, 400);

  const id = crypto.randomUUID();
  const r = await gh(env, `/repos/${env.VAULT_REPO}/actions/workflows/proposal.yml/dispatches`, {
    method: 'POST',
    body: JSON.stringify({ ref: env.VAULT_REF || 'main', inputs: { request_id: id, owner: who.login, payload: JSON.stringify(inp) } }),
  });
  if (!r.ok) {
    // 원인별 안내 (토큰 값은 응답에 넣지 않음)
    const hint = {
      401: 'GH_TOKEN 값이 틀렸거나 만료됨 — 새 토큰을 Pages 변수 GH_TOKEN 에 다시 넣고 재배포',
      403: 'GH_TOKEN 권한 부족 — Actions: Read and write 인지 확인',
      404: 'GH_TOKEN 이 proposal-vault 에 접근 못 함 또는 VAULT_REPO 값 오류 (baek05180/proposal-vault)',
      422: '워크플로 입력값·브랜치 불일치 — proposal-vault 의 proposal.yml 확인',
    }[r.status] || '';
    return json({ error: `GitHub 실행 요청 실패 (${r.status})${hint ? ' — ' + hint : ''}` }, 502);
  }
  return json({ id });
}
