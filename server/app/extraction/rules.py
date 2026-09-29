"""룰 추출 (FR-201·205): 키워드 사전(국·영문) + 정규식 + 문자열 유사도.

'라벨: 값' 줄에서 라벨을 사전과 비교해 필드를 정하고(유사도 = 점수), 값은 parsers.py로만 변환한다.
라벨 없이 문장 속에 있는 값(예: 제목의 'LCL', '40HQ*1')은 점수를 낮게 준다.
LLM이 꺼져 있어도 이 단계만으로 동작한다.
"""

import difflib
import re

from . import parsers as P
from .base import Candidate, Source, dictionary

_LABEL_LINE = re.compile(r"^\s*(?:[-•·*▪◦]|\d{1,2}[.)])?\s*([^:\n]{1,40}?)\s*:\s*(.+?)\s*$")
_TAB_LINE = re.compile(r"^\s*([^\t\n]{1,40}?)\t+(.+?)\s*$")
_REQUEST_WORDS = re.compile(r"부탁|제안|요청|확인|미정|별도|추후|문의|tbd|tba|to be|please|advise|n/?a", re.I)
_PAYMENT = re.compile(r"T\s*/?\s*T|L\s*/\s*C|D\s*/\s*[PA]|\bCAD\b|\bO/?A\b|송금|신용장|선급|후불|advance|at sight|usance|net\s*\d+", re.I)


def _norm_label(text: str) -> str:
    text = P.clean(text).lower()
    text = re.sub(r"\(.*?\)|\[.*?\]", " ", text)  # '품목(Commodity)' → '품목'
    return re.sub(r"[\s._/'\-]+", "", text)


def label_scores(label: str) -> list[tuple[str, float]]:
    """라벨 → [(필드, 유사도)] 높은 순. 완전 일치 1.0, 사전 단어 포함 0.85~0.95, 그 외 편집거리 비율."""
    raw = P.clean(label).lower()
    variants = {_norm_label(label), re.sub(r"[\s._/'\-]+", "", raw)}
    best: dict[str, float] = {}
    for field, words in dictionary()["labels"].items():
        for word in words:
            w = _norm_label(word)
            for v in variants:
                if not v or not w:
                    continue
                if v == w:
                    score = 1.0
                elif len(w) >= 2 and (v.endswith(w) or v.startswith(w)):
                    score = 0.85 + 0.1 * len(w) / len(v)
                else:
                    score = difflib.SequenceMatcher(None, v, w).ratio()
                if score > best.get(field, 0):
                    best[field] = score
    return sorted(((f, s) for f, s in best.items() if s >= 0.72), key=lambda fs: -fs[1])


def _looks_like_place(text: str) -> bool:
    return 0 < len(text) <= 30 and not _REQUEST_WORDS.search(text)


_DIGITS = re.compile(r"[\d.,/()\s~]+|약")
_UNITS_ONLY = re.compile(
    r"\b(?:ctns?|cartons?|boxe?s?|plts?|pallets?|pkgs?|packages?|pcs|ea|kgs?|lbs?|cbm|m3|cft|mm|cm|"
    r"stackable|g\.?w|n\.?w)\b|박스|개|팔레트", re.I)
_BARE_NUMBER = re.compile(rf"^(?:약\s*|approx\.?\s*|about\s*)?({P.NUM})\s*$", re.I)


def _is_real_text(value: str) -> bool:
    """숫자·단위만으로 된 값이 아닌지 ('3PKGS / 1,200KGS'는 품목명이 아니다)."""
    rest = _UNITS_ONLY.sub(" ", _DIGITS.sub(" ", value))  # 숫자를 먼저 지워야 '3PKGS'의 단위가 단어로 떨어진다
    return len(re.findall(r"[A-Za-z가-힣]", rest)) >= 2


def parse_value(field: str, value: str, label: str = "") -> tuple[list[tuple[str, object, float]], list[dict]]:
    """라벨이 정해진 값 → [(6장 필드, 값, 확신도)], notes. 파싱 실패면 빈 목록."""
    notes: list[dict] = []
    v = value.strip()
    out: list[tuple[str, object, float]] = []
    lab = P.clean(label).lower()
    if field == "commodity":
        if 1 <= len(v) <= 80 and not _REQUEST_WORDS.search(v) and _is_real_text(v):
            out.append(("commodity", v, 1.0))
    elif field == "cargoDetail":  # 'CARGO DETAIL: 3PKGS / 1,200KGS / 3.0CBM' — 한 줄에 여러 값 (발주 측 양식 표기)
        for sub in ("qty", "box", "totalCbm", "grossWeightKg"):
            out += parse_value(sub, v, label)[0]
    elif field == "qty":
        p = P.parse_qty(v)
        if p:
            qty, unit = p.value
            out += [("qty", qty, p.conf), ("qtyUnit", unit, p.conf)]
            if unit in P.PACKING_OF_UNIT:
                out.append(("packing", P.PACKING_OF_UNIT[unit], p.conf * 0.95))
    elif field == "packing":
        p = P.parse_packing(v)
        if p:
            out.append(("packing", p.value, 1.0))
    elif field == "box":
        p = P.parse_dims_mm(v)
        if p:
            out += [(k, n, p.conf) for k, n in zip(("boxL", "boxW", "boxH"), p.value)]
    elif field == "totalCbm":
        p = P.parse_cbm(v)
        bare = _BARE_NUMBER.match(P.clean(v))
        if p:
            out.append(("totalCbm", p.value, 1.0))
        elif bare and re.search(r"cbm|m3|부피|volume|measurement", lab):  # 'CBM : 1.5' — 라벨이 단위
            out.append(("totalCbm", P.to_number(bare[1]), 0.9))
    elif field == "grossWeightKg":
        p = P.parse_weight_kg(v)
        bare = _BARE_NUMBER.match(P.clean(v))
        if p:
            out.append(("grossWeightKg", p.value, 1.0))
        elif bare:  # 단위 없는 무게: 라벨에 kg가 있으면 인정, 아니면 kg/lbs 모호 → 검토 필요
            out.append(("grossWeightKg", P.to_number(bare[1]), 0.9 if re.search(r"kg|킬로", lab) else 0.6))
    elif field in ("pol", "pod"):
        code = P.port_code(v)
        if code:
            out.append((field, code, 1.0))
        elif _looks_like_place(v):
            notes.append({"field": field, "code": "PORT_UNMAPPED", "value": v, "candidates": P.port_candidates(v)})
    elif field == "pickupLocation":
        if 1 <= len(v) <= 60:
            out.append(("pickupLocation", v, 1.0))
    elif field == "cargoReadyDate":
        p = P.parse_date(v)
        if p:
            out.append(("cargoReadyDate", p.value, 1.0))
    elif field == "incoterms":
        terms = P.find_incoterms(v)
        if len(terms) == 1 and terms[0] in P.INCOTERMS:
            out.append(("incoterms", terms[0], 1.0))
        elif len(terms) > 1:
            notes.append({"field": "incoterms", "code": "MULTIPLE_INCOTERMS", "value": terms})
        elif len(terms) == 1:
            notes.append({"field": "incoterms", "code": "UNSUPPORTED_INCOTERMS", "value": terms[0]})
    elif field == "containerType":
        p = P.parse_container(v)
        if p:
            out.append(("containerType", p.value, 1.0))
            count = P.parse_container_count(v)
            if count and p.value != "LCL":
                out.append(("containerCount", count, 1.0))
    elif field == "invoiceValue":
        p = P.parse_money(v)
        if p:
            amount, ccy = p.value
            out += [("invoiceValue", amount, 1.0), ("ccy", ccy, 1.0)]
    elif field == "paymentTerm":
        if _PAYMENT.search(v) and len(v) <= 80:
            out.append(("paymentTerm", v, 1.0))
    return out, notes


def extract(sources: list[Source]) -> tuple[list[Candidate], list[dict], bool]:
    """모든 출처에서 룰 후보를 만든다. → (후보, notes, 다품목 여부)"""
    candidates: list[Candidate] = []
    notes: list[dict] = []
    multi = False
    header = re.compile(dictionary()["multiItemHeader"])

    for src in sources:
        if src.kind == "signature":
            continue
        lines = [ln for ln in src.text.split("\n") if ln.strip()]
        commodities: set[str] = set()
        section_headers = 0
        for original in lines:
            line = P.clean(original)  # 전각 콜론(：) 등 통일 — 근거(evidence)에는 원문 줄을 남긴다
            if header.match(line):
                section_headers += 1
                continue
            m = _LABEL_LINE.match(line) or _TAB_LINE.match(line)
            matched = False
            if m:
                label, value = m[1], m[2]
                for field, label_sim in label_scores(label):
                    parsed, found_notes = parse_value(field, value, label)
                    for note in found_notes:
                        notes.append({**note, "origin": src.label})
                    if parsed:
                        for name, val, conf in parsed:
                            candidates.append(Candidate(name, val, round(label_sim * conf, 3), original.strip(),
                                                        src.label, src.order, "rule"))
                            if name == "commodity":
                                commodities.add(str(val))
                        matched = True
                        break
                    if found_notes:
                        matched = True
                        break
            if not matched:
                candidates += _free_text(line, src, original)
        if section_headers >= 2 or len(commodities) >= 2:
            multi = True  # FR-210: 품목 섹션이 둘 이상이거나 품목 라벨이 서로 다른 값으로 두 번 이상
    return candidates, notes, multi


def _free_text(line: str, src: Source, original: str) -> list[Candidate]:
    """라벨 없는 문장에서 형태가 분명한 값만 낮은 점수로 (컨테이너·박스 규격·조건)."""
    out: list[Candidate] = []
    text, evidence = line.strip(), original.strip()
    p = P.parse_container(text)
    if p:
        out.append(Candidate("containerType", p.value, 0.8, evidence, src.label, src.order, "rule"))
        count = P.parse_container_count(text)
        if count and p.value != "LCL":
            out.append(Candidate("containerCount", count, 0.8, evidence, src.label, src.order, "rule"))
    p = P.parse_dims_mm(text)
    if p and p.conf == 1.0:
        for k, n in zip(("boxL", "boxW", "boxH"), p.value):
            out.append(Candidate(k, n, 0.75, evidence, src.label, src.order, "rule"))
    terms = P.find_incoterms(text)
    if len(terms) == 1 and terms[0] in P.INCOTERMS and re.search(r"조건|기준|terms?|basis|견적|quot|rfq", text, re.I):
        out.append(Candidate("incoterms", terms[0], 0.8, evidence, src.label, src.order, "rule"))
    return out
