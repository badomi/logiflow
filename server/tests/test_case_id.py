"""FR-102 케이스 ID 발급 수용 기준: LQ-YYYY-MMDD-NNN 형식, 중복 없음."""

from datetime import datetime, timezone

from app.case_id import find_case_ids, issue_case_id


def test_format_and_sequence(session):
    now = datetime(2026, 9, 28, 3, 0, tzinfo=timezone.utc)  # KST 12:00
    ids = [issue_case_id(session, now) for _ in range(3)]
    assert ids == ["LQ-2026-0928-001", "LQ-2026-0928-002", "LQ-2026-0928-003"]


def test_new_day_restarts_sequence(session):
    issue_case_id(session, datetime(2026, 9, 28, 3, 0, tzinfo=timezone.utc))
    assert issue_case_id(session, datetime(2026, 9, 29, 3, 0, tzinfo=timezone.utc)) == "LQ-2026-0929-001"


def test_date_uses_korea_time(session):
    # UTC 9/28 16:00 = KST 9/29 01:00 → 한국 날짜 기준으로 9/29
    assert issue_case_id(session, datetime(2026, 9, 28, 16, 0, tzinfo=timezone.utc)) == "LQ-2026-0929-001"


def test_sequence_survives_commit(session):
    now = datetime(2026, 9, 28, 3, 0, tzinfo=timezone.utc)
    issue_case_id(session, now)
    session.commit()
    assert issue_case_id(session, now) == "LQ-2026-0928-002"


def test_find_case_ids_in_subject():
    subject = "RE: [LQ-2026-0928-001] 보완 요청 / 참고 LQ-2026-0928-001, LQ-2026-0927-015"
    assert find_case_ids(subject) == ["LQ-2026-0928-001", "LQ-2026-0927-015"]
