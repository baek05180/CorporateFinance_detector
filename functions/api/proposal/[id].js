// GET /api/proposal/{id}            → 진행 상태 {state: queued|running|done|error}
// GET /api/proposal/{id}?download=1 → 완성된 PPT
import { json, authorize } from '../../_shared.js';

export async function onRequestGet({ request, env, params }) {
  const who = await authorize(request, env);
  if (who instanceof Response) return who;
  const id = String(params.id || '');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(id)) return json({ error: 'not found' }, 404);
  // 요청한 본인만 상태를 보고 내려받을 수 있다
  const own = await env.PROPOSALS.get(`proposals/${id}/owner.json`);
  if (!own || (await own.json()).email !== who.email) return json({ error: 'not found' }, 404);
  const st = await env.PROPOSALS.get(`proposals/${id}/status.json`);
  if (!st) return json({ state: 'queued' });
  const s = await st.json();
  if (!new URL(request.url).searchParams.has('download')) return json({ state: s.state, error: s.error, file: s.file });  // trace 는 내려보내지 않음
  if (s.state !== 'done') return json({ error: '아직 완성되지 않았습니다.' }, 409);
  const obj = await env.PROPOSALS.get(`proposals/${id}/proposal.pptx`);
  if (!obj) return json({ error: '파일이 없습니다 (7일 경과 시 삭제).' }, 404);
  const name = s.file || 'proposal.pptx';
  return new Response(obj.body, {
    headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'content-disposition': `attachment; filename="proposal.pptx"; filename*=UTF-8''${encodeURIComponent(name)}`,
      'cache-control': 'no-store',
    },
  });
}
