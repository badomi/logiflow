"""A트랙 API 수용 기준 (FR-101·102·103·104·105·207)."""

import base64
import re

from app import pipeline
from app.config import settings

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
    assert case["status"] == "수신"
    mail = case["mails"][0]
    assert mail["role"] == "ORIGINAL" and mail["hasRaw"] is True
    assert mail["snapshot"]["conversationId"] == "CONV-1"  # .eml에 없는 값은 애드인 값으로 보완
    assert [e["eventType"] for e in case["events"]] == ["CASE_CREATED", "PIPELINE_RUN"]
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
    events = [(e["eventType"], (e["detail"] or {}).get("trigger")) for e in merged["events"]]
    assert events[-2:] == [("REPLY_MERGED", None), ("PIPELINE_RUN", "REPLY_MERGED")]

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
        case = client.post("/api/cases", json=eml_body("01_lcl_request.eml")).json()["case"]
    finally:
        pipeline.register(extractor=pipeline._NotConnected())
    failed = case["events"][-1]
    assert failed["eventType"] == "PIPELINE_FAILED" and "LLM 응답 없음" in failed["detail"]["error"]
    assert client.get(f"/api/cases/{case['caseId']}").status_code == 200


def test_eml_reader_import(capsys):
    from app.eml_reader import main
    from .conftest import SAMPLES

    assert main([str(SAMPLES), "--import"]) == 0
    lines = capsys.readouterr().out.strip().splitlines()
    assert len(lines) == 6 and all(re.search(r"LQ-\d{4}-\d{4}-\d{3}", line) for line in lines)
