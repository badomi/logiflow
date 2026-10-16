"""A트랙 API 수용 기준 (FR-101·102·103·104·105·207)."""

import base64
import re
from datetime import datetime, timezone

import pytest
from sqlalchemy import func, select

from app import pipeline
from app.config import settings
from app.models import CASE_STATUSES, Case, CaseEvent, Mail, QuoteInput

from .conftest import sample


def eml_body(name: str, **snapshot) -> dict:
    body = {"emlBase64": base64.b64encode(sample(name)).decode(), "actor": "quote@leona.example.com"}
    if snapshot:
        body["snapshot"] = snapshot
    return body


def test_create_case_from_addin(client):
    """FR-101: 메일 선택 → 케이스 생성. FR-102: ID 형식. FR-103: 원문 보관."""
    res = client.post("/api/cases", json=eml_body("01_lcl_request.eml", conversationId="CONV-1"))
    assert res.status_code == 200
    data = res.json()
    assert data["duplicate"] is False
    case = data["case"]
    assert re.fullmatch(r"LQ-\d{4}-\d{4}-\d{3}", case["caseId"])
    assert case["status"] == "파싱중"  # 응답 시점에는 추출이 도는 중
    mail = case["mails"][0]
    assert mail["role"] == "ORIGINAL" and mail["hasRaw"] is True
    assert mail["snapshot"]["conversationId"] == "CONV-1"  # .eml에 없는 값은 애드인 값으로 보완
    assert [e["eventType"] for e in case["events"]] == ["CASE_CREATED", "PIPELINE_STARTED"]  # 추출은 응답 뒤에
    assert case["events"][0]["actor"] == "quote@leona.example.com"
    assert (settings.storage_dir / case["caseId"] / "mail-01" / "original.eml").read_bytes() == sample(
        "01_lcl_request.eml"
    )


def test_attachments_saved(client):
    """FR-103: 첨부파일을 분리해 저장하고 다시 열 수 있다."""
    case = client.post("/api/cases", json=eml_body("04_insurance_with_attachments.eml")).json()["case"]
    files = sorted((settings.storage_dir / case["caseId"] / "mail-01" / "attachments").iterdir())
    assert [f.name for f in files] == ["인보이스_REF-0001.pdf", "패킹리스트_REF-0001.pdf"]
    assert len(case["mails"][0]["snapshot"]["attachments"]) == 2


def test_raw_mail_and_attachments_can_be_reopened(client):
    """FR-103 수용 기준: 메일 원문과 첨부파일을 저장하고 케이스 화면에서 다시 열어볼 수 있다."""
    case = client.post("/api/cases", json=eml_body("04_insurance_with_attachments.eml")).json()["case"]
    mail = case["mails"][0]

    raw = client.get(mail["rawUrl"])
    assert raw.status_code == 200 and raw.content == sample("04_insurance_with_attachments.eml")
    assert raw.headers["content-type"].startswith("message/rfc822")

    names = [f["filename"] for f in mail["attachmentFiles"]]
    assert names == ["인보이스_REF-0001.pdf", "패킹리스트_REF-0001.pdf"]
    first = client.get(mail["attachmentFiles"][0]["url"])
    assert first.status_code == 200 and first.content.startswith(b"%PDF")

    # 다른 케이스 번호로는 열 수 없고, 없는 번호는 404
    other = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]
    att_id = mail["attachmentFiles"][0]["url"].rsplit("/", 1)[-1]
    assert client.get(f"/api/cases/{other}/attachments/{att_id}").status_code == 404
    assert client.get(f"/api/cases/{case['caseId']}/mails/9/raw").status_code == 404


def test_duplicate_mail_does_not_create_new_case(client):
    """FR-105: 같은 메일을 두 번 등록해도 케이스가 중복 생성되지 않는다."""
    first = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()
    second = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()
    assert second["duplicate"] is True
    assert second["case"]["caseId"] == first["case"]["caseId"]
    assert len(client.get("/api/cases").json()) == 1


def test_case_ids_are_sequential(client):
    a = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]
    b = client.post("/api/cases", json=eml_body("02_fob_multi_item.eml")).json()["case"]["caseId"]
    assert int(b[-3:]) == int(a[-3:]) + 1


def test_snapshot_only_input(client):
    """원문을 못 받는 환경(구버전 Outlook)에서는 패널이 읽은 값만으로도 케이스를 만든다."""
    snapshot = {
        "subject": "견적 요청",
        "from": {"name": "화주", "email": "shipper@example.com"},
        "bodyText": "LCL 견적 부탁드립니다.",
        "internetMessageId": "<only-snapshot@example.com>",
    }
    case = client.post("/api/cases", json={"snapshot": snapshot}).json()["case"]
    assert case["mails"][0]["hasRaw"] is False
    assert case["subject"] == "견적 요청"


def test_bad_input(client):
    assert client.post("/api/cases", json={}).status_code == 400
    assert client.post("/api/cases", json={"emlBase64": "***"}).status_code == 400
    assert client.get("/api/cases/LQ-2000-0101-001").status_code == 404


def test_reply_matched_by_header_then_merged(client):
    """FR-104: 회신을 헤더로 식별해 후보 제시 → 담당자 확인 후 병합. FR-207: 병합 후 재추출·재검증 실행."""
    case_id = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]

    match = client.post("/api/replies/match", json=eml_body("03_lcl_reply.eml")).json()
    assert match["alreadyRegisteredCaseId"] is None
    assert match["candidates"][0]["caseId"] == case_id
    assert match["candidates"][0]["reasons"] == ["HEADER"]

    merged = client.post(f"/api/cases/{case_id}/replies", json=eml_body("03_lcl_reply.eml")).json()
    assert [m["role"] for m in merged["mails"]] == ["ORIGINAL", "REPLY"]
    assert merged["status"] == "재파싱"  # 응답 시점: 재추출이 백그라운드에서 도는 중
    after = client.get(f"/api/cases/{case_id}").json()
    events = [(e["eventType"], (e["detail"] or {}).get("trigger")) for e in after["events"]]
    assert events[-3:] == [("REPLY_MERGED", None), ("PIPELINE_STARTED", "REPLY_MERGED"), ("PIPELINE_RUN", "REPLY_MERGED")]
    assert after["extraction"]["trigger"] == "REPLY_MERGED"

    # 병합된 뒤에는 '이미 등록됨'으로 나오고, 같은 회신을 다시 병합해도 중복되지 않는다
    again = client.post("/api/replies/match", json=eml_body("03_lcl_reply.eml")).json()
    assert again["alreadyRegisteredCaseId"] == case_id
    merged_twice = client.post(f"/api/cases/{case_id}/replies", json=eml_body("03_lcl_reply.eml")).json()
    assert len(merged_twice["mails"]) == 2


def test_reply_matched_by_subject_case_id(client):
    case_id = client.post("/api/cases", json=eml_body("02_fob_multi_item.eml")).json()["case"]["caseId"]
    snapshot = {"subject": f"RE: [{case_id}] 보완 요청 회신", "internetMessageId": "<reply-by-subject@example.com>"}
    match = client.post("/api/replies/match", json={"snapshot": snapshot}).json()
    assert match["candidates"][0] == {
        "caseId": case_id,
        "subject": "FOB 조건 국내 운송 및 선적비용 견적 문의",
        "score": 1.0,
        "reasons": ["SUBJECT_CASE_ID"],
    }


def test_compose_window_match_with_blank_fields(client):
    """작성 창 패널은 수신 시각·Message-ID가 없어 빈 값으로 보낸다 — 제목 케이스 ID로 찾아야 한다 (FR-104)."""
    case_id = client.post("/api/cases", json=eml_body("02_fob_multi_item.eml")).json()["case"]["caseId"]
    snapshot = {
        "subject": f"Re: [{case_id}] 견적 보완 요청", "from": {"name": "", "email": ""}, "to": [], "cc": [],
        "receivedAt": "", "bodyText": "", "conversationId": "", "internetMessageId": "",
        "inReplyTo": "", "references": [], "attachments": [],
    }
    response = client.post("/api/replies/match", json={"snapshot": snapshot, "emlBase64": None, "actor": "a"})
    assert response.status_code == 200
    assert response.json()["candidates"][0]["caseId"] == case_id


def test_reply_matched_by_conversation(client):
    case_id = client.post(
        "/api/cases", json=eml_body("02_fob_multi_item.eml", conversationId="CONV-FOB")
    ).json()["case"]["caseId"]
    snapshot = {"subject": "RE: 문의", "conversationId": "CONV-FOB", "internetMessageId": "<conv-reply@example.com>"}
    match = client.post("/api/replies/match", json={"snapshot": snapshot}).json()
    assert match["candidates"][0]["caseId"] == case_id
    assert match["candidates"][0]["reasons"] == ["CONVERSATION"]


def test_no_candidates_for_unrelated_mail(client):
    client.post("/api/cases", json=eml_body("01_lcl_request.eml"))
    match = client.post("/api/replies/match", json=eml_body("05_space_inquiry_html.eml")).json()
    assert match == {"alreadyRegisteredCaseId": None, "candidates": []}


def test_mail_in_other_case_cannot_be_merged(client):
    a = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]
    b = client.post("/api/cases", json=eml_body("02_fob_multi_item.eml")).json()["case"]["caseId"]
    res = client.post(f"/api/cases/{b}/replies", json=eml_body("01_lcl_request.eml"))
    assert res.status_code == 409 and a in res.json()["detail"]


def test_pipeline_failure_keeps_case(client):
    """NFR-03: 추출(B)에서 오류가 나도 케이스는 남고 이력에 실패가 기록된다."""

    class Broken:
        name = "broken"

        def extract(self, session, case):
            raise RuntimeError("LLM 응답 없음")

    pipeline.register(extractor=Broken())
    try:
        case_id = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]
    finally:
        pipeline.register(extractor=pipeline._NotConnected())
    case = client.get(f"/api/cases/{case_id}").json()
    failed = case["events"][-1]
    assert failed["eventType"] == "PIPELINE_FAILED" and "LLM 응답 없음" in failed["detail"]["error"]
    assert case["status"] == "실패"
    assert case["extraction"]["state"] == "failed" and "LLM 응답 없음" in case["extraction"]["error"]


def test_create_returns_immediately_and_extraction_runs_after(client):
    """FR-101 수용 기준: 케이스 등록 + 항목 추출 수행. 응답은 바로 오고 추출은 뒤에서 돈다."""
    first = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]
    assert first["status"] == "파싱중" and first["extraction"]["state"] == "running"

    after = client.get(f"/api/cases/{first['caseId']}").json()
    assert after["extraction"]["state"] == "not-connected"  # B트랙 추출기 연결 전
    assert after["status"] == "수신"  # 추출기가 없으면 원래 상태로 돌아온다
    assert after["extraction"]["withinLimit"] is True and after["extraction"]["elapsedMs"] >= 0


def test_extraction_result_and_60s_check(client, monkeypatch):
    """추출기가 결과를 주면 패널이 볼 수 있게 저장하고, 60초를 넘기면 withinLimit=False로 표시한다."""

    class Fake:
        name = "fake-b-track"

        def extract(self, session, case):
            return {"commodity": {"value": "스프링노트", "evidence": "품목: 스프링노트", "score": 0.95}}

    class FakeValidator:
        name = "fake-c-track"

        def validate(self, session, case, extraction):
            case.status = "정보부족"  # C트랙이 상태를 정하면 A트랙은 되돌리지 않는다
            return {"missing": ["pol", "cargoReadyDate"]}

    pipeline.register(extractor=Fake(), validator=FakeValidator())
    monkeypatch.setattr(pipeline, "TIME_LIMIT_MS", -1)  # 60초를 넘긴 상황을 흉내
    try:
        case_id = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]["caseId"]
    finally:
        pipeline.register(extractor=pipeline._NotConnected(), validator=pipeline._NotConnected())

    case = client.get(f"/api/cases/{case_id}").json()
    extraction = case["extraction"]
    assert extraction["state"] == "done"
    assert extraction["result"]["commodity"]["value"] == "스프링노트"
    assert extraction["validation"] == {"missing": ["pol", "cargoReadyDate"]}
    assert extraction["withinLimit"] is False
    assert case["status"] == "정보부족"


def test_eml_reader_import(capsys):
    from app.eml_reader import main
    from .conftest import SAMPLES

    assert main([str(SAMPLES), "--import"]) == 0
    lines = capsys.readouterr().out.strip().splitlines()
    assert len(lines) == 6 and all(re.search(r"LQ-\d{4}-\d{4}-\d{3}", line) for line in lines)


def _console_case(session, number, *, customer="테스트무역", status="수신", pol="CNSHA", pod="KRPUS", when=None):
    """D콘솔 테스트 자료: 실제 원문/요율/LLM을 사용하지 않는다."""
    case = Case(case_id=f"LQ-2026-1016-{number:03d}", status=status,
                created_at=when or datetime(2026, 10, 16, tzinfo=timezone.utc))
    if customer is not None:
        case.quote_input = QuoteInput(customerName=customer, pol=pol, pod=pod)
    case.mails.append(Mail(role="ORIGINAL", source="EML", subject=f"요청 {number}", from_email="customer@example.com"))
    session.add(case)
    return case


def test_dtrack_search_all_pages_and_unextracted_cases(client, session):
    for n in range(1, 206):
        case = _console_case(session, n, customer=None if n == 205 else "테스트무역")
        if n == 1:
            case.mails.append(Mail(role="REPLY", source="EML", subject="두 번째 메일"))
    session.commit()
    ids = []
    for page in range(1, 4):
        response = client.get("/api/dtrack/cases", params={"pageSize": 100, "page": page, "order": "asc"})
        assert response.status_code == 200
        data = response.json()
        assert data["total"] == 205 and data["statuses"] == list(CASE_STATUSES)
        ids += [c["caseId"] for c in data["items"]]
        if page == 1:
            assert data["items"][0]["subject"] == "요청 1"
        if page == 3:
            assert data["items"][-1]["customerName"] is None
    assert len(ids) == len(set(ids)) == 205
    assert ids == [f"LQ-2026-1016-{n:03d}" for n in range(1, 206)]
    clamped = client.get("/api/dtrack/cases", params={"pageSize": 100, "page": 99}).json()
    assert clamped["page"] == 3 and len(clamped["items"]) == 5
    # A의 회신 직접 선택 목록과 응답 형태는 그대로다.
    legacy = client.get("/api/cases", params={"limit": 30}).json()
    assert len(legacy) == 30
    assert set(legacy[0]) == {"caseId", "status", "subject", "fromEmail", "createdAt"}
    detail = client.get(f"/api/cases/{ids[0]}").json()
    assert len(detail["mails"]) == 2


def test_dtrack_combined_filters_port_names_and_customer_literals(client, session):
    _console_case(session, 1, customer="Alpha_100% 무역", status="정보부족")
    _console_case(session, 2, customer="AlphaX1000 무역", status="정보부족")
    _console_case(session, 3, customer="Alpha_100% 무역", status="계산완료", pod="AUSYD")
    _console_case(session, 4, customer="신규 항구", pol="ZZABC", pod="ZZXYZ")
    session.commit()
    query = {"customer": " alpha_100% ", "status": "정보부족", "pol": "상하이", "pod": "부산"}
    data = client.get("/api/dtrack/cases", params=query).json()
    assert [c["caseId"] for c in data["items"]] == ["LQ-2026-1016-001"]
    assert client.get("/api/dtrack/cases", params={"pol": "zzabc", "pod": "zzxyz"}).json()["total"] == 1
    empty = client.get("/api/dtrack/cases", params={"customer": "없는 고객", "page": 5}).json()
    assert empty["total"] == 0 and empty["page"] == 1 and empty["items"] == []


def test_dtrack_date_range_is_inclusive_kst_registration_day(client, session):
    for n, when in enumerate([
        datetime(2026, 10, 15, 14, 59, 59, tzinfo=timezone.utc),
        datetime(2026, 10, 15, 15, tzinfo=timezone.utc),
        datetime(2026, 10, 16, 14, 59, 59, 999999, tzinfo=timezone.utc),
        datetime(2026, 10, 16, 15, tzinfo=timezone.utc),
    ], 1):
        _console_case(session, n, when=when)
    session.commit()
    data = client.get("/api/dtrack/cases", params={"startDate": "2026-10-16", "endDate": "2026-10-16"}).json()
    assert [c["caseId"] for c in data["items"]] == ["LQ-2026-1016-003", "LQ-2026-1016-002"]


@pytest.mark.parametrize("sort,expected", [("customerName", [2, 1, 3]), ("status", [1, 3, 2]),
                                           ("pol", [2, 1, 3]), ("pod", [2, 1, 3])])
def test_dtrack_sort_keeps_missing_values_last(client, session, sort, expected):
    _console_case(session, 1, customer="Bravo", status="수신", pol="KRPUS", pod="USLAX")
    _console_case(session, 2, customer="Alpha", status="계산완료", pol="CNSHA", pod="KRPUS")
    _console_case(session, 3, customer=None, status="보완대기")
    session.commit()
    data = client.get("/api/dtrack/cases", params={"sort": sort, "order": "asc"}).json()
    assert [int(c["caseId"][-3:]) for c in data["items"]] == expected


@pytest.mark.parametrize("params,code", [
    ({"status": "임의 상태"}, 400), ({"startDate": "2026-10-17", "endDate": "2026-10-16"}, 400),
    ({"startDate": "날짜 아님"}, 422), ({"page": 0}, 422), ({"pageSize": 101}, 422),
    ({"sort": "drop table"}, 422), ({"order": "random"}, 422),
])
def test_dtrack_invalid_search_is_rejected(client, params, code):
    assert client.get("/api/dtrack/cases", params=params).status_code == code


def test_dtrack_read_only_and_all_statuses(client, session):
    for n, status in enumerate(CASE_STATUSES, 1):
        _console_case(session, n, status=status)
    session.commit()
    tables = [Case, QuoteInput, Mail, CaseEvent]
    before = [session.scalar(select(func.count()).select_from(t)) for t in tables]
    for status in CASE_STATUSES:
        found = client.get("/api/dtrack/cases", params={"status": status}).json()
        assert found["total"] == 1 and found["items"][0]["status"] == status
    after = [session.scalar(select(func.count()).select_from(t)) for t in tables]
    assert before == after
    assert list(session.scalars(select(Case.status).order_by(Case.id))) == list(CASE_STATUSES)
