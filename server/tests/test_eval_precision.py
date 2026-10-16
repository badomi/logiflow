"""정밀도 회귀 테스트 — 사전(fields.json)·룰을 고친 뒤 기본 평가 세트 정밀도가 떨어지면 실패한다 (NFR-01·06)."""

import pytest

from app.extraction.base import FIELDS
from eval.dataset import CASES
from eval.run_eval import evaluate, report, same


def test_rules_precision_stays_above_target():
    result = evaluate(cases=CASES)
    assert result["precision"] >= 0.90, result["errors"]  # NFR-01의 목표. 임의의 추가 합격선을 두지 않는다.
    assert result["multiAccuracy"] == 1.0
    assert result["msMax"] < 5_000  # 룰 단계는 수 초 안 (LLM 제외)


def test_model_compare_report_recommends_passing_model():
    """모델 비교 도구: 기준(정밀도·시간·오류율)을 통과한 모델을 추천하고, 실패 모델은 이유를 적는다."""
    from eval.compare import compare

    class Quiet:  # 아무것도 안 찾는 LLM → 규칙 결과만 (통과)
        def complete_json(self, *a):
            return {"fields": {}, "multipleItems": False}

    class Down:  # 매번 시간 초과
        def complete_json(self, *a):
            raise RuntimeError("LLM 응답이 45초 안에 오지 않았습니다")

    rows, text = compare({"quiet-model": Quiet(), "down-model": Down()}, cases=CASES[:6], warm=False)
    assert [r["name"] for r in rows] == ["규칙만", "quiet-model", "down-model"]
    assert "**quiet-model** 추천" in text
    assert "❌ LLM 오류 100%" in text and "⚠" in text


@pytest.mark.parametrize("got, expected, correct", [
    ("그린오피스 최서연", "그린오피스", False),  # 실제 평가에서 부분 포함으로 정답 처리되던 값
    ("최서연", "그린푸드", False),
    ("T/T 30%", "T/T 30% Advance", False),
    ("CNSHA / CNNGB", "CNSHA", False),
    ("2026-10-20 이후", "2026-10-20", False),
    (1001, 1000, False),  # 정수 수량 차이를 기존 0.5% 허용률로 가리면 안 된다.
    (15050, 15000, False),
    (0.198, 0.19, False),  # 기존 최소 0.01 오차도 작은 CBM에서는 큰 차이가 된다.
    ("  usd  ", "USD", True),
    (" 그린오피스 ", "그린오피스", True),
    (598.742, 598.742, True),  # 기존 단위 환산 정답은 유지한다.
    (500.0, 500, True),
    (True, 1, False),
    ("500", 500, False),
    (None, None, False),
])
def test_precision_compares_complete_normalized_values(got, expected, correct):
    assert same(got, expected) is correct


def test_labels_use_evidence_without_assumed_units_or_year():
    cases = {case["id"]: case for case in CASES}
    assert len(CASES) == len(cases) == 24
    assert cases["V05"]["expected"]["customerName"] == "그린푸드"
    assert "그린푸드 최서연" in cases["V05"]["mails"][0]["body"]
    assert all(cases["V11"]["expected"][f] is None for f in ("boxL", "boxW", "boxH"))
    assert cases["T06"]["expected"]["cargoReadyDate"] is None
    assert cases["T06"]["expected"].get("totalCbm") is None  # 계산한 부피를 원문 추출 정답에 섞지 않는다.
    assert all(set(case["expected"]) <= set(FIELDS) for case in CASES)


def _extracted(fields, *, multiple=False):
    result = {
        "fields": {f: {"value": None, "status": "missing", "evidence": None, "method": None, "score": 0}
                   for f in FIELDS},
        "multipleItems": multiple, "elapsedMs": 1, "llm": {"used": False},
    }
    for name, (value, status) in fields.items():
        result["fields"][name].update(value=value, status=status, evidence="평가용 원문", method="rule", score=1)
    return result


def test_precision_keeps_missing_review_and_multi_judgments_separate(monkeypatch):
    """NFR-01의 분모를 유지하면서 기존 과대 채점과 FR-210 오판을 함께 드러낸다."""
    from eval import run_eval

    case = {"id": "scoring-regression", "desc": "채점 검증", "mails": [], "expected": {
        "customerName": "그린오피스", "contactName": "최서연", "invoiceValue": 15000,
        "commodity": "노트", "qty": 10,
    }}
    result = _extracted({
        "customerName": ("그린오피스 최서연", "filled"),
        "contactName": ("최서연", "filled"),
        "invoiceValue": (15050, "filled"),
        "totalCbm": (0.38, "filled"),  # 원문에서 확정할 수 없는 값
        "qty": (10, "review"),  # 후보가 맞아도 자동으로 채운 값은 아니다.
    }, multiple=True)
    monkeypatch.setattr(run_eval.engine, "run", lambda *a, **kw: result)
    score = evaluate(cases=[case])
    assert (score["correct"], score["filled"], score["expectedTotal"]) == (1, 4, 5)
    assert score["precision"] == 1 / 4
    assert score["coverage"] == 1 / 5
    assert {e["field"] for e in score["errors"]} == {"customerName", "invoiceValue", "totalCbm"}
    assert score["multiErrors"] == [{"case": case["id"], "got": True, "expected": False}]
    assert score["multiAccuracy"] == 0
    text = report(score, "rules", None)
    assert "다품목 판정 불일치" in text and "scoring-regression | True | False" in text


@pytest.mark.parametrize("case_id, field, value", [
    ("V11", "boxL", 1200), ("V11", "boxW", 800), ("V11", "boxH", 900),
    ("T06", "cargoReadyDate", "2026-10-20"),
])
def test_unconfirmed_label_is_not_a_correct_automatic_value(monkeypatch, case_id, field, value):
    from eval import run_eval

    case = next(c for c in CASES if c["id"] == case_id)
    monkeypatch.setattr(run_eval.engine, "run", lambda *a, **kw: _extracted({field: (value, "filled")}))
    score = evaluate(cases=[case])
    assert (score["correct"], score["filled"]) == (0, 1)
    assert score["errors"][0]["field"] == field
    assert score["errors"][0]["expected"] is None
