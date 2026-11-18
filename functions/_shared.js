// 공모사채 제안서 API 공통 함수
// 환경변수(Cloudflare Pages → Settings): GH_TOKEN(secret), GH_REPO(예: baek05180/CorporateFinance_detector), GH_REF(기본 main)
// 바인딩: R2 버킷 → 변수명 PROPOSALS
export const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });

// Cloudflare Access 가 앞단에서 막지만, 설정이 빠졌을 때를 대비해 인증 헤더가 없으면 거절한다.
// (로컬 개발 등에서 끄려면 REQUIRE_ACCESS=0)
export function requireAccess(request, env) {
  if (env.REQUIRE_ACCESS === '0') return null;
  const jwt = request.headers.get('cf-access-jwt-assertion');
  const email = request.headers.get('cf-access-authenticated-user-email');
  if (!jwt || !email) return json({ error: '접근 권한이 없습니다 (Cloudflare Access 로그인 필요).' }, 401);
  return null;
}
