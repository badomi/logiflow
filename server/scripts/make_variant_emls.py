"""평가 세트의 변형 메일을 .eml 파일로 내보낸다 → samples/eml/variants/

A트랙 .eml 리더·D트랙 대량 테스트·MUST-SHIP ① 시연(샘플 20건)에 쓸 수 있다.
실행 (server 폴더에서): .venv\\Scripts\\python scripts\\make_variant_emls.py
"""

import sys
from pathlib import Path

SERVER = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SERVER))

from eval.dataset import CASES  # noqa: E402
from eval.run_eval import build_eml  # noqa: E402

OUT = SERVER / "samples" / "eml" / "variants"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    count = 0
    for case in CASES:
        generated = [m for m in case["mails"] if "eml" not in m]
        for n, mail in enumerate(generated, start=1):
            suffix = "_reply" if mail.get("role") == "REPLY" else ("" if n == 1 else f"_{n}")
            (OUT / f"{case['id']}{suffix}.eml").write_bytes(build_eml(mail))
            count += 1
    print(f"{count}개 → {OUT}")


if __name__ == "__main__":
    main()
