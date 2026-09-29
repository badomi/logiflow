"""검증: 추출된 6장 필드를 룰로 판정해 보완 문항을 만들거나, 다 갖춰졌으면 견적서를 만든다.

pipeline.py의 Validator 규격 구현. LLM을 쓰지 않는다 (정합성 검증·단가 계산은 룰 베이스 — 발주 측 승인 사항).

판정 결과 → 케이스 상태 (6장 caseStatus)
  - 다품목(FR-210) 또는 위험물 키워드 감지     → 보류   (담당자 수동 처리)
  - 화주에게 물어볼 항목이 있음               → 정보부족 (questions가 보완 요청 초안 문항이 된다, FR-304·305)
  - 전부 충족 + 요율 있음                     → 계산완료 (견적서 XLSX 생성 → A트랙 송부 초안이 첨부)
  - 전부 충족 + 요율 없음                     → 보류   (FR-501: 요율 없으면 견적을 만들지 않는다)

validation["questions"]는 docs/트랙간_전달사항.md 3-2에서 A트랙이 요청한 형식(문장 목록) 그대로다.
"""

import json
import re
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path

from sqlalchemy.orm import Session

from . import quotation
from .config import KST
from .extraction import LOCODE, port_candidates
from .models import Case, ContainerType, QuoteInput

RULES_FILE = Path(__file__).resolve().parent / "data" / "validation_rules.json"


@lru_cache
def rules() -> dict:
    """검증 룰 데이터 (필수 항목·위험물 키워드·허용 오차·화주 질문 문구). NFR-06"""
    return json.loads(RULES_FILE.read_text(encoding="utf-8"))


PACKINGS = ("Carton", "Pallet", "Crate", "Drum")
INCOTERMS = ("FOB", "CIF", "DDP", "EXW")
_UN_NUMBER = re.compile(r"\bUN\s?\d{4}\b", re.I)


def today_kst() -> date:
    return datetime.now(KST).date()


def detect_dg(commodity: str | None) -> list[str]:
    if not commodity:
        return []
    upper = commodity.upper()
    hits = [k for k in rules()["dangerousGoodsKeywords"] if k.upper() in upper]
    if _UN_NUMBER.search(commodity):
        hits.append("UN번호")
    return hits




def check(session: Session, qi: QuoteInput, notes: list[dict] | None = None, today: date | None = None,
          review: dict[str, str] | None = None, container_count: int | None = None) -> list[dict]:
    """6장 제약으로 판정한 문제 목록. question이 있는 것만 화주 문항이 되고, 나머지는 담당자 참고용.

    review: 추출은 됐지만 확신이 낮아 저장하지 않은 필드 → {필드: 근거 원문}. 이런 필드는 '맞는지 확인' 문항으로 묻는다.
    """
    today = today or today_kst()
    notes = notes or []
    review = review or {}

    def confirm(field: str) -> str:
        return (f"{rules()["labels"][field]}을(를) 「{review[field]}」 기준으로 이해했습니다. 맞는지 확인 부탁드립니다. "
                f"(Please confirm: {review[field]})")
    issues: list[dict] = []

    def add(field: str, code: str, message: str, ask: str | None = None, **extra):
        issues.append({"field": field, "code": code, "message": message, "question": ask, **extra})

    note_codes = {(n["field"], n["code"]): n for n in notes}

    for field in rules()["required"]:
        if getattr(qi, field) is not None:
            continue
        if (field, "MULTIPLE_INCOTERMS") in note_codes:
            options = " / ".join(note_codes[(field, "MULTIPLE_INCOTERMS")]["value"])
            add(field, "MULTIPLE_INCOTERMS", f"조건이 여러 개({options})",
                f"견적 조건을 {options} 중 하나로 정해 주세요. 두 조건 모두 필요하시면 말씀해 주세요. "
                f"(Please confirm one Incoterm: {options})")
        elif field in review:
            add(field, "REVIEW", "추출 신뢰도 낮음 — 검토 필요 (FR-209)", confirm(field))
        elif (field, "PORT_UNMAPPED") in note_codes:
            note = note_codes[(field, "PORT_UNMAPPED")]
            add(field, "PORT_UNMAPPED", f"'{note['value']}' UN/LOCODE 매핑 실패", rules()["questions"][field],
                candidates=note.get("candidates") or port_candidates(note["value"]))
        else:
            add(field, "MISSING", "누락", rules()["questions"][field])

    has_box = all(getattr(qi, k) is not None for k in ("boxL", "boxW", "boxH"))
    if not has_box and qi.totalCbm is None:
        if "volume" in review:
            add("volume", "REVIEW", "추출 신뢰도 낮음 — 검토 필요 (FR-209)", confirm("volume"))
        else:
            add("volume", "MISSING", "boxL/W/H·totalCbm 모두 없음", rules()["questions"]["volume"])

    if qi.qty is not None and qi.qty <= 0:
        add("qty", "INVALID", "0 초과여야 함", rules()["questions"]["qty"])
    if qi.packing is not None and qi.packing not in PACKINGS:
        add("packing", "INVALID", f"코드값 아님: {qi.packing}", rules()["questions"]["packing"])
    for side in ("pol", "pod"):
        value = getattr(qi, side)
        if value is not None and not LOCODE.match(value):
            add(side, "INVALID", f"UN/LOCODE 형식 아님: {value}", rules()["questions"][side])
    if qi.incoterms is not None and qi.incoterms not in INCOTERMS:
        add("incoterms", "INVALID", f"코드값 아님: {qi.incoterms}", rules()["questions"]["incoterms"])
    if qi.ccy is not None and not re.match(r"^[A-Z]{3}$", qi.ccy):
        add("ccy", "INVALID", f"통화코드 아님: {qi.ccy}", rules()["questions"]["ccy"])

    if qi.cargoReadyDate is not None:
        try:
            ready = date.fromisoformat(qi.cargoReadyDate)
            if ready < today:
                add("cargoReadyDate", "PAST_DATE", f"과거 날짜: {qi.cargoReadyDate}",
                    f"화물 준비일이 지난 날짜({qi.cargoReadyDate})로 되어 있습니다. 새 준비일을 알려 주세요. "
                    f"(Cargo ready date {qi.cargoReadyDate} has passed. Please advise a new date.)")
        except ValueError:
            add("cargoReadyDate", "INVALID", f"ISO 날짜 아님: {qi.cargoReadyDate}", rules()["questions"]["cargoReadyDate"])

    container = session.get(ContainerType, qi.containerType) if qi.containerType else None
    if qi.containerType is not None and container is None:
        add("containerType", "NOT_IN_MASTER", f"마스터 미등록: {qi.containerType}", rules()["questions"]["containerType"])
    # 중량 초과: 메일에 대수가 적혀 있을 때만 '1대당 중량'으로 묻는다. 대수가 없으면 견적에서 필요한 대수를 산정한다.
    if container and container.max_gross_kg and qi.grossWeightKg and container_count:
        per = qi.grossWeightKg / container_count
        if per > container.max_gross_kg:
            add("grossWeightKg", "OVER_WEIGHT",
                f"{container.code} 1대 허용중량 {container.max_gross_kg:,.0f}kg 초과 (1대당 {per:,.0f}kg)",
                f"총중량 {qi.grossWeightKg:,.0f}kg이 {container.code} {container_count}대 허용중량을 넘습니다. "
                f"컨테이너 규격·대수를 확인해 주세요. (Gross weight exceeds the limit of {container_count} x {container.code}.)")

    if has_box and qi.qty and qi.totalCbm:
        calculated = qi.boxL * qi.boxW * qi.boxH / 1e9 * qi.qty
        if calculated > 0 and abs(calculated - qi.totalCbm) / calculated > rules()["cbmTolerance"]:
            add("totalCbm", "CBM_MISMATCH", f"규격 기준 {calculated:.2f}CBM ≠ 기재 {qi.totalCbm}CBM",
                f"박스 규격·수량으로 계산한 부피는 약 {calculated:.2f}CBM인데 전체 부피는 {qi.totalCbm}CBM으로 적혀 있습니다. "
                f"정확한 부피를 확인해 주세요. (Calculated volume {calculated:.2f} CBM differs from stated {qi.totalCbm} CBM.)")
    return issues


def review_evidence(extraction: dict) -> dict[str, str]:
    """추출 결과에서 '검토 필요'(값은 있으나 임계치 미만) 필드의 근거 원문. 다품목으로 보류된 값은 제외."""
    out: dict[str, str] = {}
    for name, f in (extraction.get("fields") or {}).items():
        if f.get("status") == "review" and f.get("evidence") and f.get("reason") != "MULTIPLE_ITEMS":
            key = "volume" if name in ("boxL", "boxW", "boxH", "totalCbm") else name
            out.setdefault(key, f["evidence"].strip())
    return out


class RuleValidator:
    """pipeline.Validator 구현 — register(validator=RuleValidator())로 연결."""

    name = "rule-v1"

    def validate(self, session: Session, case: Case, extraction: dict | None) -> dict | None:
        qi = case.quote_input
        if qi is None:
            return None  # 추출 결과가 없으면 판정하지 않는다 (상태는 파이프라인이 되돌림)

        extraction = extraction or {}
        stated = (extraction.get("fields") or {}).get("containerCount") or {}
        count = stated.get("value") if stated.get("status") == "filled" else None
        issues = check(session, qi, extraction.get("notes"), review=review_evidence(extraction), container_count=count)
        questions = [i["question"] for i in issues if i["question"]]
        hold: list[dict] = []
        if extraction.get("multipleItems"):
            hold.append({"code": "MULTIPLE_ITEMS", "message": "다품목 — 수동 처리 대상 (FR-210)"})
        dg = detect_dg(qi.commodity)
        if dg and rules().get("holdOnDangerousGoods", True):
            hold.append({"code": "DANGEROUS_GOODS", "message": f"위험물 가능성 키워드: {', '.join(dg)} — 담당자 확인 필요"})

        result: dict = {"issues": issues, "questions": questions, "hold": hold, "quote": None}
        if hold:
            case.status = "보류"
        elif questions:
            case.status = "정보부족"
        else:
            quote = quotation.build(session, case, extraction)
            if quote is None:
                case.status = "보류"
                _rows, reason = quotation.rates_with_reason(session, qi, today_kst())
                hold.append({"code": "NO_RATE", "message": reason or f"요율 없음 {qi.pol}→{qi.pod} {qi.containerType} (FR-501)"})
            else:
                case.status = "계산완료"
                result["quote"] = quote
        result["status"] = case.status
        return result
