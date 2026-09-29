"""LLM 단계: 로컬 LLM이 '어디에 무슨 값이 적혀 있는지' 찾아 원문 그대로 옮긴다.

LLM에게 시키지 않는 것: 단위 환산·코드 변환·날짜 계산 (FR-203 결정형 로직 = parsers.py).
LLM이 준 값은 반드시 원문 인용(quote)과 함께 오고, grounding.py가 그 인용이 실제 메일에 있는지 검사한다.
라벨 없이 문장으로 쓴 메일("부산에서 상하이로 40피트 한 대…")처럼 룰이 약한 곳을 LLM이 메운다.
"""

from . import parsers as P
from .base import Candidate, Source

# LLM 응답 필드 → 파싱 방법. 값은 전부 '원문 그대로의 문자열'로 받는다.
LLM_FIELDS = {
    "commodity": "품목명 (예: 스프링노트, Plastic Parts)",
    "quantity": "포장 수량과 단위 원문 (예: 10박스, 500 CTNS). 낱개 수량만 있으면 그것",
    "packing": "포장 형태 원문 (예: 박스, Pallet)",
    "boxDimensions": "박스 1개 규격 원문 (예: 310 × 450 × 270mm)",
    "totalVolume": "전체 부피 원문 (예: 0.19CBM, 350 CFT)",
    "grossWeight": "총중량 원문 (예: 500kg, 2,650 LBS)",
    "portOfLoading": "선적항(출발 항구) 원문. 픽업지·공장 주소는 아님",
    "portOfDischarge": "도착항 원문",
    "pickupLocation": "픽업지·출고지 원문",
    "cargoReadyDate": "화물 준비일·출고 가능일 원문 (예: 2026-10-20, 10월 20일)",
    "incoterms": "거래 조건 원문 (예: FOB, CIF, FOB 및 CIF)",
    "container": "컨테이너 규격·수량 또는 LCL 원문 (예: 40FT HC 1대, LCL)",
    "invoiceValue": "인보이스 금액·물품 가액 원문 (예: USD 3,000)",
    "paymentTerm": "결제 조건 원문 (예: T/T 30% Advance)",
    "customerName": "요청한 고객사명 (서명의 회사명)",
    "contactName": "요청한 담당자 이름 (서명, 직함 포함 가능)",
    "contactEmail": "요청한 담당자 이메일",
}

# CPU에서도 60초 안에 끝나도록 짧게 쓴다 (출력·입력 토큰이 곧 시간)
SYSTEM_PROMPT = """견적 요청 메일에서 요청한 항목의 값을 찾아 옮겨 적어라.
- value: 메일 글자 그대로. 단위 환산·계산·번역 금지 ("2,650 LBS"는 그대로).
- quote: value가 있는 줄을 메일에서 그대로 복사.
- 메일에 없는 항목은 쓰지 말고 생략. 추측 금지.
- 뒤 메일(회신)이 앞 메일 값을 고치면 뒤 메일 값.
- 픽업지·공장 주소는 선적항이 아니다.
- 품목이 둘 이상이고 품목마다 수량·규격이 따로 있으면 multipleItems=true.
예) "총중량은 약 2,650 LBS입니다" → grossWeight {"value":"약 2,650 LBS","quote":"총중량은 약 2,650 LBS입니다"}"""

# LLM 항목 → 6장 필드. 룰이 이 필드들을 다 찾았으면 LLM에게 묻지 않는다
KEY_FIELDS = {
    "commodity": ("commodity",),
    "quantity": ("qty", "qtyUnit"),
    "packing": ("packing",),
    "boxDimensions": ("boxL", "boxW", "boxH"),
    "totalVolume": ("totalCbm",),
    "grossWeight": ("grossWeightKg",),
    "portOfLoading": ("pol",),
    "portOfDischarge": ("pod",),
    "cargoReadyDate": ("cargoReadyDate",),
    "incoterms": ("incoterms",),
    "container": ("containerType",),
    "invoiceValue": ("invoiceValue", "ccy"),
    "paymentTerm": ("paymentTerm",),
    "customerName": ("customerName",),
    "contactName": ("contactName",),
    "contactEmail": ("contactEmail",),
}
CONTACT_KEYS = ("customerName", "contactName", "contactEmail")

_ITEM = {
    "type": "object",
    "properties": {"value": {"type": ["string", "null"]}, "quote": {"type": ["string", "null"]}},
    "required": ["value", "quote"],
}
def schema_for(keys: list[str]) -> dict:
    return {
        "type": "object",
        "properties": {
            "fields": {"type": "object", "properties": {k: _ITEM for k in keys}},  # 없는 항목은 생략
            "multipleItems": {"type": "boolean"},
        },
        "required": ["fields", "multipleItems"],
    }


SCHEMA = schema_for(list(LLM_FIELDS))


def build_prompt(sources: list[Source], keys: list[str] | None = None) -> str:
    keys = keys or list(LLM_FIELDS)
    field_help = "\n".join(f"- {k}: {LLM_FIELDS[k]}" for k in keys)
    need_signature = any(k in CONTACT_KEYS for k in keys)
    blocks = "\n\n".join(
        f"[{s.label}]\n{s.text}" for s in sources
        if s.text.strip() and (s.kind != "signature" or need_signature)
    )
    return f"찾을 항목:\n{field_help}\n\n=== 메일 ===\n{blocks}"


def to_candidates(raw: dict, sources: list[Source], ground,
                  asked: list[str] | None = None) -> tuple[list[Candidate], list[dict], bool]:
    """LLM 원시 응답 → 후보. 값 변환은 parsers.py, 점수는 grounding(인용 검증). 묻지 않은 항목은 무시."""
    fields = raw.get("fields") or {}
    if asked is not None:
        fields = {k: v for k, v in fields.items() if k in asked}
    out: list[Candidate] = []
    notes: list[dict] = []

    def add(name: str, value, conf: float, item: dict):
        score, origin, order = ground(item.get("quote"), item.get("value"), sources)
        if score > 0:
            out.append(Candidate(name, value, round(score * conf, 3), item.get("quote") or "", origin, order, "llm"))

    for key, item in fields.items():
        if not isinstance(item, dict) or not item.get("value"):
            continue
        text = str(item["value"]).strip()
        if key == "commodity":
            add("commodity", text, 1.0, item)
        elif key == "quantity":
            p = P.parse_qty(text)
            if p:
                add("qty", p.value[0], p.conf, item)
                add("qtyUnit", p.value[1], p.conf, item)
                if p.value[1] in P.PACKING_OF_UNIT:
                    add("packing", P.PACKING_OF_UNIT[p.value[1]], p.conf * 0.95, item)
        elif key == "packing":
            p = P.parse_packing(text)
            if p:
                add("packing", p.value, 1.0, item)
        elif key == "boxDimensions":
            p = P.parse_dims_mm(text)
            if p:
                for name, n in zip(("boxL", "boxW", "boxH"), p.value):
                    add(name, n, p.conf, item)
        elif key == "totalVolume":
            p = P.parse_cbm(text)
            if p:
                add("totalCbm", p.value, 1.0, item)
        elif key == "grossWeight":
            p = P.parse_weight_kg(text)
            if p:
                add("grossWeightKg", p.value, 1.0, item)
        elif key in ("portOfLoading", "portOfDischarge"):
            name = "pol" if key == "portOfLoading" else "pod"
            code = P.port_code(text)
            if code:
                add(name, code, 1.0, item)
            else:
                notes.append({"field": name, "code": "PORT_UNMAPPED", "value": text, "candidates": P.port_candidates(text)})
        elif key == "pickupLocation":
            add("pickupLocation", text, 1.0, item)
        elif key == "cargoReadyDate":
            p = P.parse_date(text)
            if p:
                add("cargoReadyDate", p.value, 1.0, item)
        elif key == "incoterms":
            terms = P.find_incoterms(text)
            if len(terms) == 1 and terms[0] in P.INCOTERMS:
                add("incoterms", terms[0], 1.0, item)
            elif len(terms) > 1:
                notes.append({"field": "incoterms", "code": "MULTIPLE_INCOTERMS", "value": terms})
        elif key == "container":
            p = P.parse_container(text)
            if p:
                add("containerType", p.value, 1.0, item)
                count = P.parse_container_count(text)
                if count and p.value != "LCL":
                    add("containerCount", count, 1.0, item)
        elif key == "invoiceValue":
            p = P.parse_money(text)
            if p:
                add("invoiceValue", p.value[0], 1.0, item)
                add("ccy", p.value[1], 1.0, item)
        elif key in ("paymentTerm", "customerName", "contactName", "contactEmail"):
            add(key, text, 1.0, item)
    return out, notes, bool(raw.get("multipleItems"))
