"""docs/proposal_spec.md 의 '## 설정' yaml 블록을 읽는다 (제안서 생성기·사이트 팝업 공용)."""
import os
import re

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPEC_PATH = os.path.join(ROOT, "docs", "proposal_spec.md")


def load_spec(path=SPEC_PATH):
    text = open(path, encoding="utf-8").read()
    part = text.split("## 설정", 1)[-1]
    m = re.search(r"```yaml\s*\n(.*?)```", part, re.S)
    if not m:
        raise ValueError("proposal_spec.md 에 '## 설정' yaml 블록이 없습니다")
    spec = yaml.safe_load(m.group(1)) or {}
    # yaml 이 날짜 키를 date 로 바꾸므로 문자열로 통일
    for k in ("events", "holidays"):
        spec[k] = {str(d): v for d, v in (spec.get(k) or {}).items()}
    return spec


if __name__ == "__main__":
    import json
    print(json.dumps(load_spec(), ensure_ascii=False, indent=1)[:2000])
