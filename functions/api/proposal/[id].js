// GET /api/proposal/{id}            → {state: queued|running|done|error}
// GET /api/proposal/{id}?download=1 → 완성된 PPT
// 실행 이름이 "proposal {id} by {아이디}" 인 비공개 저장소 실행을 찾고, 요청한 본인일 때만 응답한다.
import { authorize, gh, json } from '../../_shared.js';
import { unzipEntry } from '../../_zip.js';

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

async function findRun(env, id) {
  const r = await gh(env, `/repos/${env.VAULT_REPO}/actions/workflows/proposal.yml/runs?event=workflow_dispatch&per_page=40`);
  if (!r.ok) throw new Error(`runs ${r.status}`);
  return ((await r.json()).workflow_runs || []).find(x => String(x.display_title || x.name || '').startsWith(`proposal ${id} `));
}

async function artifactZip(env, run, id) {
  const a = await gh(env, `/repos/${env.VAULT_REPO}/actions/runs/${run.id}/artifacts?name=proposal-${id}`).then(r => r.json());
  const art = (a.artifacts || [])[0];
  if (!art || art.expired) return null;
  // GitHub 은 임시 저장소 주소로 넘겨준다. GitHub 토큰이 그쪽으로 따라가지 않도록 직접 따라간다.
  const r = await gh(env, `/repos/${env.VAULT_REPO}/actions/artifacts/${art.id}/zip`, { redirect: 'manual' });
  const loc = r.headers.get('location');
  const z = loc ? await fetch(loc) : r;
  if (!z.ok) throw new Error(`artifact ${z.status}`);
  return z.arrayBuffer();
}

export async function onRequestGet({ request, env, params }) {
  const who = await authorize(request, env);
  if (who instanceof Response) return who;
  const id = String(params.id || '');
  if (!UUID.test(id)) return json({ error: 'not found' }, 404);
  let run;
  try { run = await findRun(env, id); } catch { return json({ error: 'GitHub 조회 실패' }, 502); }
  if (!run) return json({ state: 'queued' });  // 실행 직후 몇 초간은 목록에 안 보일 수 있음
  if (!String(run.display_title || '').endsWith(` by ${who.login}`)) return json({ error: 'not found' }, 404);
  if (run.status !== 'completed') return json({ state: run.status === 'in_progress' ? 'running' : 'queued' });

  const zip = await artifactZip(env, run, id).catch(() => null);
  if (!zip) return json({ state: 'error', error: run.conclusion === 'success' ? '결과 파일이 없습니다 (7일 경과 시 삭제).' : '생성 실패 — 비공개 저장소의 Actions 기록을 확인해 주세요.' });
  const st = JSON.parse(new TextDecoder().decode((await unzipEntry(zip, 'status.json')) || new TextEncoder().encode('{}')));
  if (!new URL(request.url).searchParams.has('download')) return json({ state: st.state || 'error', error: st.error, file: st.file });
  const pptx = st.state === 'done' ? await unzipEntry(zip, 'proposal.pptx') : null;
  if (!pptx) return json({ error: '아직 완성되지 않았거나 파일이 없습니다.' }, 409);
  return new Response(pptx, {
    headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'content-disposition': `attachment; filename="proposal.pptx"; filename*=UTF-8''${encodeURIComponent(st.file || 'proposal.pptx')}`,
      'cache-control': 'no-store',
    },
  });
}
