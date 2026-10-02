"""LLM 단계 (LLM 중심 추출): 로컬 LLM이 메일의 모든 항목을 읽어 '정리된 값 + 근거 줄 번호'로 답한다.

설계
- 메일 줄마다 번호(L1, L2 …)를 붙여 보내고, LLM은 값과 **줄 번호**만 돌려준다 → 근거 문장을 베끼지 않아 출력이 짧고
  (CPU 노트북에서도 전 항목 추출이 60초 안에), 근거는 실제 원문 줄이라 지어낼 수 없다 (FR-202).
- LLM은 '뜻'을 정리한다: 말로 쓴 수(오백, 반 톤, 한 대)는 숫자로, 다른 줄에 나뉜 금액·통화는 합쳐서, 번호로만 답한 회신은
  직전 보완 질문에 맞춰서. **계산·단위 환산·코드 변환은 하지 않는다** → 규칙(parsers.py, FR-203 결정형)이 한다.
- LLM 값은 규칙과 같은 해석기(rules.parse_value)를 거친다. 해석하지 못하면 버리지 않고 '확인 필요'로 남긴다.
- 값의 숫자가 근거 줄에 없으면(말로 쓴 수도 아니면) 지어낸 값으로 보고 버린다 (grounding.support).
"""

import re

from .base import GROUP_FIELDS, UNPARSED_SCORE, Candidate, Source

# LLM에게 묻는 항목. 값은 '숫자 단위'처럼 정리된 문자열로 받는다 (환산은 규칙이 한다)
LLM_FIELDS = {
    "commodity": "품목명 (적힌 그대로, 번역하지 않음)",
    "quantity": "포장 수량과 단위 (예: 10 CTNS, 5 pallets). 낱개 수량만 있으면 그것",
    "packing": "포장 형태 (carton / pallet / crate / drum 등 적힌 말)",
    "boxDimensions": "박스 1개 규격 '가로 x 세로 x 높이 단위' (예: 310 x 450 x 270 mm)",
    "totalVolume": "전체 부피와 단위 (예: 0.38 CBM, 350 CFT)",
    "grossWeight": "총중량과 적힌 단위 그대로 (예: 500 kg, 2650 lbs, 0.5 ton)",
    "portOfLoading": "선적 항구 이름 (픽업지·공장 주소 아님)",
    "portOfDischarge": "도착 항구 이름",
    "cargoReadyDate": "화물 준비일 YYYY-MM-DD (연도가 안 적혀 있으면 '10월 20일'처럼 적힌 그대로)",
    "incoterms": "거래 조건 (FOB, CIF 등). 여러 개면 모두",
    "container": "컨테이너 규격과 대수 (예: 40HQ x 2) 또는 LCL",
    "invoiceValue": "인보이스 금액과 통화 (예: 8500 USD). 금액과 통화가 다른 줄이면 합쳐서",
    "currency": "인보이스 통화만 적힌 경우 (예: USD)",
    "paymentTerm": "결제 조건 (예: T/T 30% Advance)",
    "carrier": "고객이 지정한 선사 (예: HMM). 지정이 없으면 생략",
    "customerName": "견적을 요청한 회사명",
    "contactName": "견적을 요청한 담당자 이름",
    "contactEmail": "견적을 요청한 담당자 이메일",
}

SYSTEM_PROMPT = """견적 요청 메일(줄마다 L번호)에서 요청한 항목의 값을 찾아라.
- v: 값. 단위는 적힌 그대로 '숫자 단위'로 정리 (계산·단위 환산 금지). 말로 쓴 수(오백, 반, 한 대)는 숫자로.
- l: 값이 있는 줄 번호 목록. 금액과 통화처럼 두 줄에 나뉘면 둘 다 (예: [5, 6]).
- 메일에 없는 항목은 쓰지 말고 생략. 추측 금지.
- 뒤 메일(회신)이 앞 메일 값을 고치면 뒤 메일 값.
- 회신이 '1. 인천'처럼 번호로만 답하면 '직전 질문'의 번호에 맞춰 해석.
- 픽업지·공장 주소는 선적항이 아니다. 선적항을 정해 달라는 요청이면 portOfLoading은 생략.
- 품목이 둘 이상이고 품목마다 수량·규격이 따로 있으면 multipleItems=true.
예) L3 "무게는 오백 킬로 정도요" → grossWeight {"v":"500 kg","l":[3]}
예) L5 "인보이스 금액: 8,500" L6 "통화: USD" → invoiceValue {"v":"8500 USD","l":[5,6]}"""

# LLM 항목 → 6장 필드 (빈 칸만 묻는 모드에서 어떤 항목을 물을지 정할 때)
KEY_FIELDS = {
    "commodity": ("commodity",), "quantity": ("qty", "qtyUnit"), "packing": ("packing",),
    "boxDimensions": ("boxL", "boxW", "boxH"), "totalVolume": ("totalCbm",), "grossWeight": ("grossWeightKg",),
    "portOfLoading": ("pol",), "portOfDischarge": ("pod",), "cargoReadyDate": ("cargoReadyDate",),
    "incoterms": ("incoterms",), "container": ("containerType",), "invoiceValue": ("invoiceValue",),
    "currency": ("ccy",), "paymentTerm": ("paymentTerm",), "customerName": ("customerName",),
    "contactName": ("contactName",), "contactEmail": ("contactEmail",), "carrier": ("carrier",),
}
CONTACT_KEYS = ("customerName", "contactName", "contactEmail")

# LLM 항목 → 규칙 해석기(rules.parse_value)의 라벨 그룹. 해석은 규칙과 한 곳에서만 한다
KEY_GROUP = {
    "commodity": "commodity", "quantity": "qty", "packing": "packing", "boxDimensions": "box",
    "totalVolume": "totalCbm", "grossWeight": "grossWeightKg", "portOfLoading": "pol", "portOfDischarge": "pod",
    "cargoReadyDate": "cargoReadyDate", "incoterms": "incoterms",
    "container": "containerType", "invoiceValue": "invoiceValue", "currency": "ccy", "paymentTerm": "paymentTerm",
    "carrier": "carrier",
}
# 보완 질문 항목 → 프롬프트에 보여 줄 이름 (번호 답변 해석용)
FIELD_NAMES = {
    "commodity": "품목", "qty": "수량", "qtyUnit": "수량 단위", "packing": "포장 형태", "volume": "박스 규격 또는 전체 부피",
    "totalCbm": "전체 부피", "pol": "선적항", "pod": "도착항", "cargoReadyDate": "화물 준비일", "incoterms": "거래 조건",
    "containerType": "컨테이너", "grossWeightKg": "총중량", "invoiceValue": "인보이스 금액·통화", "ccy": "통화",
    "paymentTerm": "결제 조건",
}

_ITEM = {
    "type": "object",
    "properties": {"v": {"type": "string"}, "l": {"type": "array", "items": {"type": "integer"}}},
    "required": ["v", "l"],
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


def build_prompt(sources: list[Source], keys: list[str] | None = None,
                 asked: list[str | None] | None = None) -> tuple[str, list[tuple[Source, str]]]:
    """메일에 줄 번호를 붙인 프롬프트와, 줄 번호 → (출처, 원문 줄) 목록."""
    keys = keys or list(LLM_FIELDS)
    need_signature = any(k in CONTACT_KEYS for k in keys)
    index: list[tuple[Source, str]] = []
    blocks = []
    for src in sources:
        if not src.text.strip() or (src.kind == "signature" and not need_signature):
            continue
        lines = [ln for ln in src.text.split("\n") if ln.strip()]
        numbered = []
        for ln in lines:
            index.append((src, ln.strip()))
            numbered.append(f"L{len(index)}: {ln.strip()}")
        blocks.append(f"[{src.label}]\n" + "\n".join(numbered))
    field_help = "\n".join(f"- {k}: {LLM_FIELDS[k]}" for k in keys)
    prompt = f"찾을 항목:\n{field_help}\n"
    if asked:
        prompt += "\n직전 질문 (화주가 번호로 답했을 수 있음):\n" + "\n".join(
            f"{n}) {FIELD_NAMES.get(f or '', '(담당자가 쓴 질문)')}" for n, f in enumerate(asked, start=1)) + "\n"
    prompt += "\n=== 메일 ===\n" + "\n\n".join(blocks)
    return prompt, index


_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
from .rules import _PAYMENT  # noqa: E402  결제 방식을 뜻하는 말 (티티, 송금, 신용장, 선결제 …)


def to_candidates(raw: dict, index: list[tuple[Source, str]], sources: list[Source] | None = None,
                  asked_keys: list[str] | None = None) -> tuple[list[Candidate], list[dict], bool]:
    """LLM 응답 → 후보. 근거는 줄 번호의 원문, 해석은 rules.parse_value, 해석 실패는 '확인 필요'."""
    from . import grounding
    from .rules import parse_value

    fields = raw.get("fields") or {}
    if asked_keys is not None:
        fields = {k: v for k, v in fields.items() if k in asked_keys}
    out: list[Candidate] = []
    notes: list[dict] = []

    for key, item in fields.items():
        if not isinstance(item, dict):
            continue
        text = str(item.get("v") or item.get("value") or "").strip()
        if not text:
            continue
        if "l" in item:
            support, src, evidence = grounding.support_lines(text, item.get("l") or [], index)
        else:  # 예전 형식 {value, quote} (테스트·구버전 호환)
            support, label, order = grounding.ground(item.get("quote"), text, sources or [s for s, _ in index])
            src = next((s for s in (sources or []) if s.label == label), None)
            evidence = item.get("quote") or ""
            if src is None and support > 0:
                src = Source(order, label, "body", evidence)
        group = KEY_GROUP.get(key)
        if src is not None and group and support < 1.0:
            same, text = _same_meaning(group, text, evidence)  # 하이큐브=40HQ, 싱가포르=Singapore, 10월 20일=2026-10-20
            if same:
                support = 1.0
        if support <= 0 or src is None:
            continue  # 근거 줄이 없거나 값의 숫자가 근거에 없음 → 지어낸 값으로 보고 버린다
        if key == "paymentTerm" and not _PAYMENT.search(evidence):
            support = round(support * 0.6, 3)  # 원문 줄에 결제 방식을 뜻하는 말이 없으면 저장하지 않는다(확인 필요)

        if key in ("customerName", "contactName"):
            out.append(Candidate(key, text, support, evidence, src.label, src.order, "llm"))
            continue
        if key == "contactEmail":
            if _EMAIL.match(text):
                out.append(Candidate(key, text, support, evidence, src.label, src.order, "llm"))
            continue
        if group is None:
            continue
        parsed, found = parse_value(group, text, LLM_FIELDS.get(key, key))
        notes += found
        for name, value, conf in parsed:
            out.append(Candidate(name, value, round(support * conf, 3), evidence, src.label, src.order, "llm"))
        if not parsed and not found and not _other_kind(group, text):
            # 못 읽었지만 다른 종류의 값도 아님 → 담당자 확인용으로 남긴다 (화주 질문에 인용하지는 않음)
            for name in GROUP_FIELDS.get(group, ()):
                out.append(Candidate(name, text, UNPARSED_SCORE, evidence, src.label, src.order, "llm", unparsed=True))
    return out, notes, bool(raw.get("multipleItems"))


# 값의 종류를 알아보는 해석기 — 인보이스 금액 칸에 '360 kg'처럼 다른 종류가 오면 LLM이 줄을 잘못 짚은 것
def _other_kind(group: str, text: str) -> bool:
    from . import parsers as P

    kinds = {
        "grossWeightKg": P.parse_weight_kg(text), "totalCbm": P.parse_cbm(text), "box": P.parse_dims_mm(text),
        "invoiceValue": P.parse_money(text), "cargoReadyDate": P.parse_date(text),
        "containerType": P.parse_container(text), "pol": P.port_code(text), "qty": P.parse_qty(text),
    }
    same_family = {"pol": {"pol", "pod"}, "pod": {"pol", "pod"}, "qty": {"qty", "qtyUnit"}, "ccy": {"invoiceValue"}}
    mine = same_family.get(group, {group})
    return any(found and kind not in mine for kind, found in kinds.items())


def _same_meaning(group: str, value: str, evidence: str) -> tuple[bool, str]:
    """LLM이 정리한 값과 근거 원문 줄이 규칙 해석기로 읽었을 때 같은 값인가 → (같음, 쓸 값).

    날짜는 원문에 연도가 없으면 LLM이 붙인 연도를 믿지 않고 원문 그대로('10월 20일') 넘겨 규칙이 추정하게 한다.
    """
    from . import parsers as P
    from .rules import parse_value

    if group in ("pol", "pod"):
        code = P.port_code(value)
        return bool(code and code in P.ports_in(evidence)), value
    mine, _ = parse_value(group, value, group)
    theirs, _ = parse_value(group, re.sub(r"^[^:：]{1,30}[:：]\s*", "", evidence), group)
    if not mine or not theirs:
        return False, value
    same = mine[0][1] == theirs[0][1]
    if same and group == "cargoReadyDate" and theirs[0][2] < 1.0:
        return True, re.sub(r"^[^:：]{1,30}[:：]\s*", "", evidence)  # 연도는 규칙이 추정 (확신도 0.8)
    return same, value
