"""GitHub Actions(proposal.yml)에서 호출: 입력 JSON 파일 → out/proposal.pptx + out/status.json

공개 저장소의 Actions 기록은 누구나 볼 수 있으므로, 입력값·회사명·오류 상세를 로그에 찍지 않는다.
오류 상세는 결과물(status.json, 비공개 저장소 7일 보관)에만 남긴다.
"""
import contextlib
import io
import json
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tools import proposal5, proposal6  # noqa: E402
from tools.proposal_spec import load_spec  # noqa: E402


def main(template, inp_path, outdir):
    os.makedirs(outdir, exist_ok=True)
    buf = io.StringIO()
    try:
        with open(inp_path, encoding="utf-8") as f:
            inp = json.load(f)
        out = os.path.join(outdir, "proposal.pptx")
        deck = str((load_spec().get("deck") or "v5")).lower()
        checks = []
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            if deck in ("v5", "v5_basic"):
                proposal5.run(template, inp, out, mix=(deck == "v5"))
                checks = list(proposal5.LAST_CHECKS)
            else:
                _, checks = proposal6.run(template, inp, out)
        co = {}
        p = os.path.join("public", "data", "co", f"{inp['corp']}.json")
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                co = json.load(f)
        name = (co.get("name") or inp["corp"]).replace("/", "_")
        kind = "회사채_발행검토자료" if deck == "v6" else "공모사채_발행제안서"
        st = {"state": "done", "file": f"{name}_{kind}_{inp['issue']}.pptx", "checks": checks}
    except Exception as e:  # noqa: BLE001
        st = {"state": "error", "error": f"{type(e).__name__}: {str(e)[:300]}", "trace": traceback.format_exc()[-3000:]}
    with open(os.path.join(outdir, "status.json"), "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False)
    print("result:", st["state"])  # 로그에는 성공/실패만
    return st["state"] == "done"


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1], sys.argv[2], sys.argv[3]) else 1)
