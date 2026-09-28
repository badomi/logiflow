"""케이스 ID 발급 (FR-102): LQ-YYYY-MMDD-NNN, 날짜(KST)별 순번, 중복 없음."""

import re
from datetime import datetime, timezone

from sqlalchemy.orm import Session

from .config import KST
from .models import CaseSeq

# 메일 제목 등에서 케이스 ID를 찾는 패턴 (회신 식별 FR-104에서 사용)
CASE_ID_PATTERN = re.compile(r"LQ-\d{4}-\d{4}-\d{3}")

MAX_SEQ_PER_DAY = 999


def issue_case_id(session: Session, now: datetime | None = None) -> str:
    """오늘 날짜의 다음 순번으로 케이스 ID를 만든다.

    순번 행을 잠그고(with_for_update) 1을 올리므로, 동시에 여러 명이 눌러도 같은 번호가 나오지 않는다.
    호출한 쪽의 트랜잭션이 commit될 때 순번이 확정된다.
    """
    local = (now or datetime.now(timezone.utc)).astimezone(KST)
    date_key = local.strftime("%Y%m%d")

    row = session.get(CaseSeq, date_key, with_for_update=True)
    if row is None:
        row = CaseSeq(date_key=date_key, last_seq=0)
        session.add(row)
    if row.last_seq >= MAX_SEQ_PER_DAY:
        raise RuntimeError(f"{date_key} 케이스 ID 순번이 {MAX_SEQ_PER_DAY}을 넘었습니다.")
    row.last_seq += 1
    session.flush()

    return f"LQ-{local:%Y}-{local:%m%d}-{row.last_seq:03d}"


def find_case_ids(text: str) -> list[str]:
    """문자열 안의 케이스 ID를 순서대로 중복 없이 돌려준다."""
    return list(dict.fromkeys(CASE_ID_PATTERN.findall(text or "")))
