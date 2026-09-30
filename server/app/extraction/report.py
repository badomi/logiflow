"""놓친 표현 모아 보기 — 추출이 값을 찾았지만 해석하지 못한 원문(UNPARSED)과 '확인 필요' 값을 항목별로 모은다.

    python -m app.extraction.report            최근 30일
    python -m app.extraction.report --days 7
    GET /api/extraction/review-report?days=30

관리자가 보고 자주 나오는 표현을 사전에 추가하면(app/extraction/data/fields.json 라벨, parsers.py 단위·포장 단어)
다음부터는 자동으로 읽는다. 사전을 고친 뒤에는 pytest와 python -m eval.run_eval로 정밀도가 떨어지지 않았는지 확인.
"""

import argparse
import json
from collections import defaultdict
from datetime import datetime, timedelta, timezone

from sqlalchemy import select
from sqlalchemy.orm import Session

from ..models import Case, ExtractionField


def collect(session: Session, days: int = 30, examples: int = 5) -> dict:
    since = datetime.now(timezone.utc) - timedelta(days=days)
    rows = session.execute(
        select(ExtractionField, Case.case_id).join(Case, Case.id == ExtractionField.case_pk)
        .where(ExtractionField.created_at >= since, ExtractionField.status == "review")
        .order_by(ExtractionField.id.desc())
    ).all()
    unparsed: dict[str, dict] = defaultdict(lambda: {"count": 0, "examples": []})
    low: dict[str, dict] = defaultdict(lambda: {"count": 0, "examples": []})
    seen: set[tuple] = set()
    for row, case_id in rows:
        key = (case_id, row.field_name, row.evidence)
        if key in seen:  # 같은 케이스를 여러 번 추출해도 한 번만 센다
            continue
        seen.add(key)
        bucket = unparsed if (row.method or "").endswith(":unparsed") else low
        entry = bucket[row.field_name]
        entry["count"] += 1
        if len(entry["examples"]) < examples:
            entry["examples"].append({"case": case_id, "evidence": row.evidence,
                                      "value": json.loads(row.value) if row.value else None})
    order = lambda d: dict(sorted(d.items(), key=lambda kv: -kv[1]["count"]))  # noqa: E731
    return {"days": days, "unparsed": order(unparsed), "lowConfidence": order(low)}


def main() -> None:
    from ..db import SessionLocal, init_db

    parser = argparse.ArgumentParser(description="놓친 표현 모아 보기")
    parser.add_argument("--days", type=int, default=30)
    args = parser.parse_args()
    init_db()
    with SessionLocal() as session:
        report = collect(session, args.days)
    print(f"최근 {report['days']}일 — 값을 찾았지만 읽지 못한 표현 (사전 보강 후보)")
    if not report["unparsed"]:
        print("  없음")
    for field, info in report["unparsed"].items():
        print(f"  [{field}] {info['count']}건")
        for ex in info["examples"]:
            print(f"      {ex['case']}: {ex['evidence']}")
    print()
    print("확신이 낮아 저장하지 않은 값 (담당자 확인)")
    if not report["lowConfidence"]:
        print("  없음")
    for field, info in report["lowConfidence"].items():
        print(f"  [{field}] {info['count']}건 — 예: {info['examples'][0]['evidence']}")


if __name__ == "__main__":
    main()
