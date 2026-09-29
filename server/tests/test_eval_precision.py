"""정밀도 회귀 테스트 — 사전(fields.json)·룰을 고친 뒤 기본 평가 세트 정밀도가 떨어지면 실패한다 (NFR-01·06)."""

from eval.dataset import CASES
from eval.run_eval import evaluate


def test_rules_precision_stays_above_target():
    result = evaluate(cases=CASES)
    assert result["precision"] >= 0.95, result["errors"]  # 목표 90% + 여유. 떨어지면 오추출 목록을 확인
    assert result["multiAccuracy"] == 1.0
    assert result["msMax"] < 5_000  # 룰 단계는 수 초 안 (LLM 제외)
