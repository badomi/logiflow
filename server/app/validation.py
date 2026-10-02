"""검증: 추출된 6장 필드를 룰로 판정해 보완 문항을 만들거나, 다 갖춰졌으면 견적서를 만든다.

pipeline.py의 Validator 규격 구현. LLM을 쓰지 않는다 (정합성 검증·단가 계산은 룰 베이스 — 발주 측 승인 사항).

판정 결과 → 케이스 상태 (6장 caseStatus)
  - 다품목(FR-210) 또는 위험물·냉동·특수화물 키워드 감지(FR-308) → 보류   (담당자 수동 처리)
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


def josa(word: str, with_final: str, without_final: str) -> str:
    """받침에 맞는 조사: 총중량+을, 포장 형태+를. 한글이 아닌 글자로 끝나면(360kg) 받침 있는 쪽."""
    last = word.strip()[-1:] if word.strip() else ""
    if "가" <= last <= "힣":
        return with_final if (ord(last) - 0xAC00) % 28 else without_final
    return with_final


SEVERITIES = ("Block", "Warn", "Info")


def severity(code: str) -> str:
    """문제 코드 → 심각도 Block/Warn/Info (FR-301). validation_rules.json의 "severity"로 바꾼다 (NFR-06).

    정의서 2장: Block은 견적 산출 불가, Warn은 산출 후 사람 검수 필요, Info는 참고.
    화주 문항이나 보류 사유가 있으면 견적을 만들지 않으므로 그런 코드는 Block이어야 한다 (Warn·Info로 두지 말 것).
    """
    table = rules().get("severity") or {}
    level = table.get(code) or table.get("default") or "Block"
    return level if level in SEVERITIES else "Block"


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




def detect_special(commodity: str | None) -> list[str]:
    """냉동·냉장·초과규격 등 특수화물 키워드 (FR-308). 영문 키워드는 단어 단위로만 본다 (OOG가 GOOGLE에 걸리지 않게)."""
    if not commodity:
        return []
    upper = commodity.upper()
    hits = []
    for keyword in rules().get("specialCargoKeywords") or []:
        k = keyword.upper()
        if k.isascii():
            if re.search(rf"(?<![A-Z]){re.escape(k)}(?![A-Z])", upper):
                hits.append(keyword)
        elif k in upper:
            hits.append(keyword)
    return hits


def check(session: Session, qi: QuoteInput, notes: list[dict] | None = None, today: date | None = None,
          review: dict[str, str] | None = None, container_count: int | None = None) -> list[dict]:
    """6장 제약으로 판정한 문제 목록. question이 있는 것만 화주 문항이 되고, 나머지는 담당자 참고용.

    review: 추출은 됐지만 저장하지 않은 필드 → {필드: {"evidence": 근거 원문, "reason": 사유}}.
      확신이 낮으면 '맞는지 확인', 해석하지 못했으면(UNPARSED) 원문을 짚어 정확한 형식으로 다시 묻는다.
    """
    today = today or today_kst()
    notes = notes or []
    review = review or {}

    def confirm(field: str) -> str:
        info = review[field]
        label, ev = rules()["labels"][field], info["evidence"]
        if info.get("reason") == "UNPARSED":
            hint = rules().get("formatHints", {}).get(field)
            example = f" ({hint} 형식)" if hint else ""
            return (f"{label}{josa(label, '을', '를')} 「{ev}」{josa(ev, '이라고', '라고')} 적어 주셨는데 정확히 읽지 못했습니다. "
                    f"다시 알려 주세요{example}. (Please re-confirm: {ev})")
        return (f"{label}{josa(label, '을', '를')} 「{ev}」 기준으로 이해했습니다. 맞는지 확인 부탁드립니다. "
                f"(Please confirm: {ev})")
    issues: list[dict] = []

    def add(field: str, code: str, message: str, ask: str | None = None, **extra):
        issues.append({"field": field, "code": code, "severity": severity(code), "message": message,
                       "question": ask, **extra})

    note_codes = {(n["field"], n["code"]): n for n in notes}

    for field in rules()["required"]:
        if getattr(qi, field) is not None:
            continue
        if (field, "MULTIPLE_INCOTERMS") in note_codes:
            options = " / ".join(note_codes[(field, "MULTIPLE_INCOTERMS")]["value"])
            add(field, "MULTIPLE_INCOTERMS", f"조건이 여러 개({options})",
                f"견적 조건을 {options} 중 하나로 정해 주세요. 두 조건 모두 필요하시면 말씀해 주세요. "
                f"(Please confirm one Incoterm: {options})")
        elif (field, "UNSUPPORTED_INCOTERMS") in note_codes:
            term = note_codes[(field, "UNSUPPORTED_INCOTERMS")]["value"]
            add(field, "UNSUPPORTED_INCOTERMS", f"지원하지 않는 조건: {term}",
                f"요청하신 {term} 조건은 현재 FOB·CIF·DDP·EXW 기준으로 견적드리고 있습니다. 어느 조건으로 견적할지 알려 주세요. "
                f"(We currently quote on FOB/CIF/DDP/EXW. Which term should we use instead of {term}?)")
        elif field in review:
            code = "UNPARSED" if review[field].get("reason") == "UNPARSED" else "REVIEW"
            add(field, code, "값을 찾았지만 해석 못 함" if code == "UNPARSED" else "추출 신뢰도 낮음 — 검토 필요 (FR-209)",
                confirm(field))
        elif (field, "PORT_UNMAPPED") in note_codes:
            note = note_codes[(field, "PORT_UNMAPPED")]
            add(field, "PORT_UNMAPPED", f"'{note['value']}' UN/LOCODE 매핑 실패", rules()["questions"][field],
                candidates=note.get("candidates") or port_candidates(note["value"]))
        else:
            add(field, "MISSING", "누락", rules()["questions"][field])

    # 금액과 통화가 둘 다 없으면 '금액과 통화' 질문 하나만 (통화 질문이 따로 또 나가지 않게)
    if qi.invoiceValue is None and qi.ccy is None:
        issues[:] = [i for i in issues if i["field"] != "ccy"]

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
    # 허용중량 자료가 없는 컨테이너(예: 40RF)는 초과 검사를 못 한다 → 담당자 참고용으로만 알린다 (화주 문항 아님)
    if container and container.code != "LCL" and not container.max_gross_kg and qi.grossWeightKg:
        add("grossWeightKg", "NO_WEIGHT_LIMIT",
            f"{container.code} 허용중량 자료가 없어 중량 초과 검사를 하지 않았습니다 (총중량 {qi.grossWeightKg:,.0f}kg)")
    # 부피 초과: 중량과 같이 메일에 대수가 적혀 있을 때만 묻는다. 대수가 없으면 견적에서 필요한 대수를 산정한다.
    cbm = qi.totalCbm
    if cbm is None and has_box and qi.qty:
        cbm = qi.boxL * qi.boxW * qi.boxH / 1e9 * qi.qty
    if container and container.max_cbm and cbm and container_count and cbm > container.max_cbm * container_count:
        add("totalCbm", "OVER_VOLUME",
            f"{container.code} {container_count}대 용적 {container.max_cbm * container_count:,.0f}CBM 초과 (화물 {cbm:,.2f}CBM)",
            f"화물 부피 {cbm:,.2f}CBM이 {container.code} {container_count}대에 실을 수 있는 부피를 넘습니다. "
            f"컨테이너 규격·대수를 확인해 주세요. (Cargo volume exceeds the capacity of {container_count} x {container.code}.)")

    if has_box and qi.qty and qi.totalCbm:
        calculated = qi.boxL * qi.boxW * qi.boxH / 1e9 * qi.qty
        if calculated > 0 and abs(calculated - qi.totalCbm) / calculated > rules()["cbmTolerance"]:
            add("totalCbm", "CBM_MISMATCH", f"규격 기준 {calculated:.2f}CBM ≠ 기재 {qi.totalCbm}CBM",
                f"박스 규격·수량으로 계산한 부피는 약 {calculated:.2f}CBM인데 전체 부피는 {qi.totalCbm}CBM으로 적혀 있습니다. "
                f"정확한 부피를 확인해 주세요. (Calculated volume {calculated:.2f} CBM differs from stated {qi.totalCbm} CBM.)")
    return issues


def review_evidence(extraction: dict) -> dict[str, dict]:
    """추출 결과에서 '확인 필요'(저장 안 한 값) 필드의 근거 원문과 사유. 다품목으로 보류된 값은 제외."""
    out: dict[str, dict] = {}
    for name, f in (extraction.get("fields") or {}).items():
        if f.get("reason") == "UNPARSED" and (f.get("method") or "").startswith("llm"):
            continue  # LLM이 짚은 줄은 틀릴 수 있다 → 원문 인용 없이 일반 질문으로 (담당자 화면에는 '확인 필요'로 보임)
        if f.get("status") == "review" and f.get("evidence") and f.get("reason") != "MULTIPLE_ITEMS":
            key = "volume" if name in ("boxL", "boxW", "boxH", "totalCbm") else name
            out.setdefault(key, {"evidence": f["evidence"].strip(), "reason": f.get("reason")})
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
            hold.append({"code": "MULTIPLE_ITEMS", "severity": severity("MULTIPLE_ITEMS"), "message": "다품목 — 수동 처리 대상 (FR-210)"})
        dg = detect_dg(qi.commodity)
        if dg and rules().get("holdOnDangerousGoods", True):
            hold.append({"code": "DANGEROUS_GOODS", "severity": severity("DANGEROUS_GOODS"), "message": f"위험물 가능성 키워드: {', '.join(dg)} — 담당자 확인 필요"})
        special = detect_special(qi.commodity)
        # 냉동·특수 컨테이너(40RF 등)로 요청한 건도 자동 견적 범위 밖이다 (정의서 3장 범위 제외, FR-308)
        if qi.containerType in (rules().get("specialContainers") or []):
            special.append(f"{qi.containerType} 컨테이너")
        if special and rules().get("holdOnSpecialCargo", True):
            hold.append({"code": "SPECIAL_CARGO", "severity": severity("SPECIAL_CARGO"),
                         "message": f"냉동·특수화물 가능성: {', '.join(special)} — 담당자 처리 (FR-308)"})

        result: dict = {"issues": issues, "questions": questions, "hold": hold, "quote": None}
        if hold:
            case.status = "보류"
        elif questions:
            case.status = "정보부족"
        else:
            quote = quotation.build(session, case, extraction)
            if quote is None:
                case.status = "보류"
                _rows, reason = quotation.rates_with_reason(session, qi, quotation.rate_date(qi, today_kst()))
                hold.append({"code": "NO_RATE", "severity": severity("NO_RATE"), "message": reason or f"요율 없음 {qi.pol}→{qi.pod} {qi.containerType} (FR-501)"})
            else:
                case.status = "계산완료"
                result["quote"] = quote
        result["status"] = case.status
        levels = [i["severity"] for i in issues + hold]
        result["severityCount"] = {level: levels.count(level) for level in SEVERITIES}
        result["questionCount"] = len(questions)  # 보완 요청 초안의 문항 수가 이 값과 같아야 한다 (FR-305)
        return result
