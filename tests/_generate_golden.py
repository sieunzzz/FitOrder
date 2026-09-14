"""현재 동작을 golden 파일로 저장한다. 동작을 의도적으로 바꾼 뒤에만 실행한다.

    python tests/_generate_golden.py                 # 전체 재생성
    python tests/_generate_golden.py ledger ui_merges # 일부만 재생성
"""
from __future__ import annotations

import json
import sys

from _cases import GOLDEN_PATH, SNAPSHOTS


def main(names):
    data = json.loads(GOLDEN_PATH.read_text(encoding="utf-8")) if GOLDEN_PATH.exists() else {}
    for name in names or SNAPSHOTS:
        data[name] = SNAPSHOTS[name]()
        print(f"updated: {name}")
    GOLDEN_PATH.parent.mkdir(parents=True, exist_ok=True)
    GOLDEN_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main(sys.argv[1:])
