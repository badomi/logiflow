"""정밀도 회귀 테스트 — 사전(fields.json)·룰을 고친 뒤 기본 평가 세트 정밀도가 떨어지면 실패한다 (NFR-01·06)."""

from eval.dataset import CASES
from eval.run_eval import evaluate


def test_rules_precision_stays_above_target():
    result = evaluate(cases=CASES)
    assert result["precision"] >= 0.95, result["errors"]  # 목표 90% + 여유. 떨어지면 오추출 목록을 확인
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
