// 공모사채 제안서 API 공통 함수
//
// Cloudflare Pages → Settings 에 넣는 값
//   GH_TOKEN (Secret)      GitHub fine-grained 토큰 — 이 저장소 Actions: Read and write 만
//   GH_REPO                예: baek05180/CorporateFinance_detector
//   GH_REF                 기본 main
//   ACCESS_TEAM_DOMAIN     예: myteam.cloudflareaccess.com (Zero Trust → Settings → Custom pages 의 팀 도메인)
//   ACCESS_AUD             Access 애플리케이션의 Application Audience (AUD) Tag
//   ALLOWED_EMAILS         쓸 수 있는 사람 (쉼표 구분). 비우면 Access 정책을 통과한 사람 모두
// 바인딩: R2 버킷 → 변수명 PROPOSALS
export const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });

const b64url = s => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(s.length / 4) * 4, '=')), c => c.charCodeAt(0));
const dec = s => JSON.parse(new TextDecoder().decode(b64url(s)));

let CERTS = { at: 0, keys: [] };
async function certs(team) {
  if (Date.now() - CERTS.at < 3600e3 && CERTS.keys.length) return CERTS.keys;
  const r = await fetch(`https://${team}/cdn-cgi/access/certs`);
  if (!r.ok) throw new Error('certs');
  CERTS = { at: Date.now(), keys: (await r.json()).keys || [] };
  return CERTS.keys;
}

// Cloudflare Access 가 발급한 토큰(JWT)의 서명·대상·만료를 직접 검증한다.
// 헤더만 보는 방식은 Access 를 거치지 않는 주소(*.pages.dev, 미리보기 배포)에서 위조할 수 있기 때문.
// 성공하면 { email }, 실패하면 Response(401/403) 를 돌려준다.
export async function authorize(request, env) {
  const deny = (m, s = 401) => json({ error: m }, s);
  if (!env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD) return deny('서버 설정이 끝나지 않았습니다 (ACCESS_TEAM_DOMAIN·ACCESS_AUD).', 500);
  const tok = request.headers.get('cf-access-jwt-assertion') || (request.headers.get('cookie') || '').match(/CF_Authorization=([^;]+)/)?.[1];
  if (!tok) return deny('로그인이 필요합니다 (Cloudflare Access).');
  try {
    const [h, p, sig] = tok.split('.');
    const head = dec(h), claims = dec(p);
    const jwk = (await certs(env.ACCESS_TEAM_DOMAIN)).find(k => k.kid === head.kid);
    if (!jwk || head.alg !== 'RS256') return deny('인증 토큰이 올바르지 않습니다.');
    const key = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
    const ok = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, b64url(sig), new TextEncoder().encode(`${h}.${p}`));
    const aud = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
    const now = Date.now() / 1000;
    if (!ok || !aud.includes(env.ACCESS_AUD) || claims.iss !== `https://${env.ACCESS_TEAM_DOMAIN}` || !(claims.exp > now) || (claims.nbf && claims.nbf > now + 60))
      return deny('인증 토큰이 올바르지 않거나 만료되었습니다.');
    const email = String(claims.email || '').toLowerCase();
    const allow = String(env.ALLOWED_EMAILS || '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
    if (!email || (allow.length && !allow.includes(email))) return deny('사용 권한이 없는 계정입니다.', 403);
    return { email };
  } catch {
    return deny('인증 확인에 실패했습니다.');
  }
}
