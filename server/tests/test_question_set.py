import json

import pytest
from sqlalchemy import select

from app.db import SessionLocal
from app.models import Case, CaseEvent
from .test_drafts import make_case, set_validation

QUESTIONS = ["박스 규격 또는 총부피(CBM)를 알려 주세요.", "결제 조건(T/T·L/C)을 알려 주세요."]
FIELDS = ["volume", "paymentTerm"]


def ready(client):
    case_id = make_case(client)
    set_validation(case_id, QUESTIONS, FIELDS)
    return case_id


def post(client, case_id, questions):
    return client.post(f"/api/cases/{case_id}/drafts/supplement", json={"questions": questions})


@pytest.mark.parametrize("questions", [QUESTIONS[:1], QUESTIONS + ["선적항을 알려 주세요."],
    [QUESTIONS[0], QUESTIONS[0]], [QUESTIONS[0], "인보이스 금액을 알려 주세요."],
    [QUESTIONS[0], "결제 조건(T/T·L/C)과 전화번호를 알려 주세요."]])
def test_mismatch_rejected_without_draft_event(client, questions):
    case_id = ready(client)
    response = post(client, case_id, questions)
    assert response.status_code == 400
    assert not any(e["eventType"] == "DRAFT_PREPARED" for e in client.get(f"/api/cases/{case_id}").json()["events"])


def test_reorder_and_polite_edit_keep_reply_field_mapping(client):
    case_id = ready(client)
    edited = [QUESTIONS[1].replace("알려 주세요", "회신 부탁드립니다"), QUESTIONS[0]]
    response = post(client, case_id, edited)
    assert response.status_code == 200
    assert response.json()["htmlBody"].count("<li>") == 2
    event = client.get(f"/api/cases/{case_id}").json()["events"][-1]["detail"]
    assert event["questionFields"] == ["paymentTerm", "volume"]
    assert event["questions"] == edited
    assert event["questionSetVerified"] and event["validationEventId"]


def test_blank_uses_current_validation(client):
    case_id = ready(client)
    assert post(client, case_id, [" ", ""]).status_code == 200
    assert client.get(f"/api/cases/{case_id}").json()["events"][-1]["detail"]["questions"] == QUESTIONS


@pytest.mark.parametrize("event_type", ["PIPELINE_STARTED", "PIPELINE_FAILED"])
def test_stale_validation_is_not_used(client, event_type):
    case_id = ready(client)
    with SessionLocal() as session:
        case = session.scalar(select(Case).where(Case.case_id == case_id))
        session.add(CaseEvent(case=case, event_type=event_type, detail="{}"))
        session.commit()
    assert post(client, case_id, QUESTIONS).status_code == 400


def test_no_validation_or_no_missing_fields_disallows_manual_question(client):
    case_id = make_case(client)
    assert post(client, case_id, QUESTIONS).status_code == 400
    set_validation(case_id, [], [])
    assert post(client, case_id, QUESTIONS).status_code == 400


def test_old_questions_rejected_after_new_validation(client):
    case_id = ready(client)
    set_validation(case_id, [QUESTIONS[0], "선적항을 알려 주세요."], ["volume", "pol"])
    assert post(client, case_id, QUESTIONS).status_code == 400


def test_combined_invoice_question_counts_as_one(client):
    case_id = make_case(client)
    question = "인보이스 금액과 통화를 알려 주세요."
    set_validation(case_id, [question], ["invoiceValue"])
    response = post(client, case_id, [question])
    assert response.status_code == 200 and response.json()["htmlBody"].count("<li>") == 1


def test_condition_numbers_must_not_change(client):
    case_id = make_case(client)
    set_validation(case_id, ["기재한 1.5CBM이 맞는지 확인 부탁드립니다."], ["totalCbm"])
    assert post(client, case_id, ["기재한 15CBM이 맞는지 확인 부탁드립니다."]).status_code == 400
