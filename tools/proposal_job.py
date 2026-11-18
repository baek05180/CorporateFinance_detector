"""GitHub Actions(proposal.yml)에서 호출: 환경변수 PAYLOAD(JSON) → out/proposal.pptx + out/status.json"""
import json
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tools import proposal5  # noqa: E402


def main(template, outdir):
    os.makedirs(outdir, exist_ok=True)
    st = {"state": "error"}
    try:
        inp = json.loads(os.environ["PAYLOAD"])
        out = os.path.join(outdir, "proposal.pptx")
        proposal5.run(template, inp, out)
        co = {}
        p = os.path.join("public", "data", "co", f"{inp['corp']}.json")
        if os.path.exists(p):
            co = json.load(open(p, encoding="utf-8"))
        name = (co.get("name") or inp["corp"]).replace("/", "_")
        st = {"state": "done", "file": f"{name}_공모사채_발행제안서_{inp['issue']}.pptx"}
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        st = {"state": "error", "error": f"{type(e).__name__}: {str(e)[:300]}"}
    with open(os.path.join(outdir, "status.json"), "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False)
    return st["state"] == "done"


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1], sys.argv[2]) else 1)
