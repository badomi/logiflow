"""C트랙 자동 문항의 국·영문 병기와 기존 초안·번호 회신 계약을 검증한다."""

from datetime import date
from html import escape
import re
from types import SimpleNamespace as NS

import pytest

from app import validation
from app.models import QuoteInput
from .test_pipeline_flow import connect, create, eml


TODAY = date(2026, 10, 16)
CATALOG = NS(get=lambda model, code: NS(code=code, max_gross_kg=20000, max_cbm=25))


def valid_input(**changes):
    values = dict(commodity="Notebooks", qty=10, qtyUnit="CTNS", packing="Carton",
                  boxL=1000, boxW=1000, boxH=500, totalCbm=5, pol="CNSHA", pod="KRPUS",
                  cargoReadyDate="2026-10-20", incoterms="CIF", containerType="20GP",
                  grossWeightKg=500, invoiceValue=15000, ccy="USD", paymentTerm="T/T Advance")
    return QuoteInput(**(values | changes))


def assert_bilingual_single_question(question):
    assert re.search("[가-힣]", question)
    assert "Please " in question
    assert "\n" not in question and "\r" not in question  # 패널의 split("\n")이 질문을 늘리지 않는다


@pytest.mark.parametrize("field,english", [
    ("commodity", "commodity name"), ("qty", "cargo quantity"), ("qtyUnit", "quantity unit"),
    ("packing", "packing type"), ("volume", "dimensions of one box in mm"),
    ("pol", "port of loading"), ("pod", "port of discharge"), ("cargoReadyDate", "YYYY-MM-DD"),
    ("incoterms", "Incoterms"), ("containerType", "shipment is LCL"), ("grossWeightKg", "gross weight"),
    ("invoiceValue", "invoice value and currency"), ("ccy", "invoice currency"),
    ("paymentTerm", "advance payment percentage"),
])
def test_missing_field_generates_one_complete_bilingual_question(field, english):
    qi = valid_input()
    if field == "volume":
        qi.boxL = qi.boxW = qi.boxH = qi.totalCbm = None
    else:
        setattr(qi, field, None)
    issues = validation.check(CATALOG, qi, today=TODAY)
    questions = [i for i in issues if i["question"]]
    assert len(questions) == 1 and questions[0]["field"] == field
    question = questions[0]["question"]
    assert_bilingual_single_question(question)
    assert english in question


@pytest.mark.parametrize("code,changes,notes,review,english", [
    ("REVIEW", {}, [], {"grossWeightKg": {"evidence": "총중량 500kg", "reason": "REVIEW"}},
     ['gross weight', '"총중량 500kg"', "Please confirm whether this is correct."]),
    ("UNPARSED", {"grossWeightKg": None}, [],
     {"grossWeightKg": {"evidence": "총중량\n오백 킬로", "reason": "UNPARSED"}},
     ["could not read the gross weight", '"총중량 오백 킬로"', "Please provide it again (e.g. 500kg)."]),
    ("MULTIPLE_INCOTERMS", {"incoterms": None},
     [{"field": "incoterms", "code": "MULTIPLE_INCOTERMS", "value": ["FOB", "CIF"]}], {},
     ["Please select one Incoterm from FOB / CIF.", "quotations for both"]),
    ("UNSUPPORTED_INCOTERMS", {"incoterms": None},
     [{"field": "incoterms", "code": "UNSUPPORTED_INCOTERMS", "value": "FCA"}], {},
     ["FOB/CIF/DDP/EXW", "Please specify which term to use instead of FCA."]),
    ("PAST_DATE", {"cargoReadyDate": "2026-10-01"}, [], {},
     ["2026-10-01 has passed", "Please advise a new date."]),
    ("OVER_WEIGHT", {"grossWeightKg": 50000}, [], {},
     ["50,000kg", "1 x 20GP", "Please confirm the container type and quantity."]),
    ("OVER_VOLUME", {"totalCbm": 100, "boxL": None, "boxW": None, "boxH": None}, [], {},
     ["100.00 CBM", "1 x 20GP", "Please confirm the container type and quantity."]),
    ("CBM_MISMATCH", {"totalCbm": 6}, [], {},
     ["5.00 CBM", "6 CBM", "Please confirm the correct volume."]),
])
def test_dynamic_questions_keep_values_and_requests_in_both_languages(code, changes, notes, review, english):
    issues = validation.check(CATALOG, valid_input(**changes), notes=notes, review=review,
                              today=TODAY, container_count=1)
    questions = [i for i in issues if i["question"]]
    assert len(questions) == 1 and questions[0]["code"] == code
    question = questions[0]["question"]
    assert_bilingual_single_question(question)
    assert all(part in question for part in english)


def test_bilingual_questions_reach_existing_draft_without_extra_numbers(client, connect):
    case = create(client, eml("01_lcl_request.eml"))
    result = case["extraction"]["validation"]
    questions = result["questions"]
    # A 패널이 자동 문항을 표시한 뒤 그대로 보내는 경로를 재현한다.
    submitted = "\n".join(questions).split("\n")
    assert submitted == questions
    draft = client.post(f"/api/cases/{case['caseId']}/drafts/supplement",
                        json={"questions": submitted}).json()
    assert draft["questionCheck"]["matches"] is True
    assert draft["htmlBody"].count("<li>") == result["questionCount"] == len(questions)
    for question in questions:
        assert_bilingual_single_question(question)
        assert f"<li>{escape(question)}</li>" in draft["htmlBody"]


def test_operator_question_is_preserved_without_translation(client, connect):
    case = create(client, eml("01_lcl_request.eml"))
    manual = "담당자가 직접 쓴 질문 <원문>을 확인해 주세요."
    draft = client.post(f"/api/cases/{case['caseId']}/drafts/supplement",
                        json={"questions": [manual]}).json()
    assert re.findall(r"<li>(.*?)</li>", draft["htmlBody"]) == [escape(manual)]
    assert draft["questionCheck"]["added"] == [manual]
