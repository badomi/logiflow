"""초안 내용 수용 기준 (FR-304·305·505·510). 발송 기능은 없다 — 초안 내용만 만든다."""

import base64

from .conftest import sample


def make_case(client) -> str:
    body = {"emlBase64": base64.b64encode(sample("01_lcl_request.eml")).decode(), "actor": "quote@leona.example.com"}
    return client.post("/api/cases", json=body).json()["case"]["caseId"]


def test_supplement_draft(client):
    case_id = make_case(client)
    questions = ["선적항(POL)을 알려 주세요.", "화물 준비일(Cargo Ready Date)을 알려 주세요."]
    draft = client.post(
        f"/api/cases/{case_id}/drafts/supplement", json={"questions": questions, "actor": "quote@leona.example.com"}
    ).json()

    assert draft["kind"] == "SUPPLEMENT"
    assert draft["subject"].startswith(f"[{case_id}]")  # FR-304: 제목에 케이스 ID
    assert draft["to"] == [{"name": "화주 담당자", "email": "shipper@example.com"}]
    assert draft["htmlBody"].count("<li>") == len(questions)  # FR-305: 문항 수 = 누락 항목 수
    assert "안녕하세요" in draft["htmlBody"] and "Hello" in draft["htmlBody"]  # FR-510: 국·영문 병기

    events = client.get(f"/api/cases/{case_id}").json()["events"]
    assert events[-1]["eventType"] == "DRAFT_PREPARED" and events[-1]["detail"]["questionCount"] == 2


def test_pending_draft_until_saved(client):
    """작성 창 버튼 보조 경로: 방금 준비한 초안을 돌려주고, 초안함 저장이 기록되면 더 이상 돌려주지 않는다."""
    case_id = make_case(client)
    actor = "quote@leona.example.com"
    assert client.get("/api/drafts/pending", params={"actor": actor}).json() is None

    client.post(f"/api/cases/{case_id}/drafts/supplement", json={"questions": ["POL?"], "actor": actor})
    pending = client.get("/api/drafts/pending", params={"actor": actor}).json()
    assert pending["caseId"] == case_id and pending["kind"] == "SUPPLEMENT"
    assert pending["subject"].startswith(f"[{case_id}]")
    assert client.get("/api/drafts/pending", params={"actor": "other@example.com"}).json() is None

    client.post(
        f"/api/cases/{case_id}/mail-events",
        json={"eventType": "DRAFT_SAVED", "subject": pending["subject"], "actor": actor},
    )
    assert client.get("/api/drafts/pending", params={"actor": actor}).json() is None


def test_supplement_draft_escapes_html(client):
    case_id = make_case(client)
    draft = client.post(f"/api/cases/{case_id}/drafts/supplement", json={"questions": ["<script>x</script>"]}).json()
    assert "<script>" not in draft["htmlBody"] and "&lt;script&gt;" in draft["htmlBody"]


def test_supplement_draft_needs_questions(client):
    case_id = make_case(client)
    res = client.post(f"/api/cases/{case_id}/drafts/supplement", json={"questions": ["  "]})
    assert res.status_code == 400


def test_sent_event_records_time_actor_and_status(client):
    """FR-505: 발송 시각·수행자 이력. 보완 요청을 보내면 '보완대기', 견적서를 보내면 '발송완료'."""
    case_id = make_case(client)
    actor = "quote@leona.example.com"

    saved = client.post(
        f"/api/cases/{case_id}/mail-events",
        json={"eventType": "DRAFT_SAVED", "subject": f"RE: [{case_id}] 견적 보완 요청", "actor": actor},
    ).json()
    assert saved["status"] == "수신"  # 초안 저장만으로는 상태가 바뀌지 않는다

    sent = client.post(
        f"/api/cases/{case_id}/mail-events",
        json={"eventType": "MAIL_SENT", "subject": f"RE: [{case_id}] 견적 보완 요청", "actor": actor},
    ).json()
    assert sent["status"] == "보완대기"
    last = sent["events"][-1]
    assert last["eventType"] == "MAIL_SENT" and last["actor"] == actor and last["createdAt"]
    assert last["detail"] == {
        "kind": "SUPPLEMENT",
        "subject": f"RE: [{case_id}] 견적 보완 요청",
        "statusFrom": "수신",
        "statusTo": "보완대기",
    }

    quote = client.post(
        f"/api/cases/{case_id}/mail-events",
        json={"eventType": "MAIL_SENT", "kind": "QUOTE", "subject": "x", "actor": actor},
    ).json()
    assert quote["status"] == "발송완료"


def test_sent_mail_recorded_once_with_real_send_time(client):
    """보낸 편지함의 메일로 기록하면 실제 발송 시각을 남기고, 같은 메일은 한 번만 기록한다."""
    case_id = make_case(client)
    body = {
        "eventType": "MAIL_SENT",
        "subject": f"[{case_id}] 견적서 송부 / Quotation",
        "actor": "quote@leona.example.com",
        "internetMessageId": "<sent-1@leona.example.com>",
        "occurredAt": "2026-09-28T13:05:00+09:00",
    }
    first = client.post(f"/api/cases/{case_id}/mail-events", json=body).json()
    second = client.post(f"/api/cases/{case_id}/mail-events", json=body).json()
    sent = [e for e in second["events"] if e["eventType"] == "MAIL_SENT"]
    assert len(sent) == 1
    assert sent[0]["detail"]["sentAt"] == "2026-09-28T04:05:00+00:00"
    assert first["status"] == second["status"] == "발송완료"


def test_sent_time_taken_from_eml_date_header(client):
    """보낸 메일 원문을 함께 보내면 Date 헤더(실제 발송 시각)를 쓴다 — Outlook 생성 시각(초안 작성 시각)이 아니라."""
    case_id = make_case(client)
    body = {
        "eventType": "MAIL_SENT",
        "subject": f"RE: [{case_id}] 견적 보완 요청",
        "internetMessageId": "<sample-03-lcl-reply@example.com>",
        "occurredAt": "2026-09-28T09:00:00+09:00",  # 틀린 값(초안 생성 시각)을 줘도
        "emlBase64": base64.b64encode(sample("03_lcl_reply.eml")).decode(),  # 원문 Date: 2026-09-29 09:00 KST
    }
    events = client.post(f"/api/cases/{case_id}/mail-events", json=body).json()["events"]
    assert events[-1]["detail"]["sentAt"] == "2026-09-29T00:00:00+00:00"


def test_leona_sent_mail_is_not_a_reply(client):
    """발송 기록된 LEONA 메일(자기 사본)은 회신 후보로 나오지 않고 병합도 막힌다 — 실제 테스트에서 발견한 오병합 방지."""
    case_id = make_case(client)
    sent_id = "<leona-sent-1@leona.example.com>"
    client.post(
        f"/api/cases/{case_id}/mail-events",
        json={"eventType": "MAIL_SENT", "subject": f"RE: [{case_id}] 견적 보완 요청", "internetMessageId": sent_id},
    )
    copy = {"snapshot": {"subject": f"RE: [{case_id}] 견적 보완 요청", "internetMessageId": sent_id}}
    assert client.post("/api/replies/match", json=copy).json()["candidates"] == []
    assert client.post(f"/api/cases/{case_id}/replies", json=copy).status_code == 409

    reply = {"snapshot": {"subject": f"Re: [{case_id}] 견적 보완 요청", "internetMessageId": "<customer-reply@x.com>"}}
    assert client.post("/api/replies/match", json=reply).json()["candidates"][0]["caseId"] == case_id


def test_unknown_mail_event_rejected(client):
    case_id = make_case(client)
    res = client.post(f"/api/cases/{case_id}/mail-events", json={"eventType": "AUTO_SEND"})
    assert res.status_code == 400


def test_quote_draft_attaches_quote_files(client):
    case_id = make_case(client)
    assert client.post(f"/api/cases/{case_id}/drafts/quote", json={}).status_code == 400  # 견적서 없음

    pdf = b"%PDF-1.4 sample quote"
    up = client.post(f"/api/cases/{case_id}/quotes", files={"file": ("견적서.pdf", pdf, "application/pdf")})
    assert up.status_code == 200
    assert client.post(
        f"/api/cases/{case_id}/quotes", files={"file": ("memo.txt", b"x", "text/plain")}
    ).status_code == 400

    draft = client.post(f"/api/cases/{case_id}/drafts/quote", json={"actor": "quote@leona.example.com"}).json()
    assert draft["kind"] == "QUOTE" and draft["subject"].startswith(f"[{case_id}]")
    assert [a["filename"] for a in draft["attachments"]] == ["견적서.pdf"]

    downloaded = client.get(draft["attachments"][0]["url"])
    assert downloaded.status_code == 200 and downloaded.content == pdf
    assert client.get(f"/api/cases/{case_id}/quotes/..%2F..%2Fx.pdf").status_code == 404
