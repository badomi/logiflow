"""결정형 파서 (FR-203): 원문 문자열 → 6장 형식 값. LLM은 여기에 관여하지 않는다.

LBS→KGS, CFT→CBM, inch/cm→mm 환산, 날짜 ISO 변환, 금액·통화, 수량·포장 단위, 인코텀즈·컨테이너 코드,
항구명 → UN/LOCODE. 모두 같은 입력이면 항상 같은 결과가 나온다.
정밀도 우선: 애매하면 None (추측 금지). 예) 연도 없는 날짜, 단위 없는 박스 규격은 낮은 신뢰도.
"""

import difflib
import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from functools import lru_cache
from pathlib import Path

DATA = Path(__file__).resolve().parent / "data"

NUM = r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?"


def clean(text: str) -> str:
    """전각 문자·특수 곱하기 기호 통일 (：→:, ×·ｘ→x 등). 비교·파싱 전에 항상 거친다."""
    text = unicodedata.normalize("NFKC", text or "")
    return text.replace("×", "x").replace("✕", "x").replace("＊", "*").replace("\u00a0", " ")


def to_number(text: str) -> float:
    return float(text.replace(",", ""))


# ------------------------------------------------------------------ 무게·부피·길이


@dataclass
class Parsed:
    value: object
    conf: float  # 파싱 확신도 0~1 (단위가 명시되면 1.0)
    raw: str


_WEIGHT = {"kg": 1, "kgs": 1, "kilo": 1, "kilos": 1, "킬로": 1, "킬로그램": 1, "g": 0.001,
           "t": 1000, "ton": 1000, "tons": 1000, "mt": 1000, "tonne": 1000, "톤": 1000,
           "lb": 0.45359237, "lbs": 0.45359237, "파운드": 0.45359237}
_VOLUME = {"cbm": 1, "m3": 1, "m³": 1, "㎥": 1, "입방미터": 1, "루베": 1,
           "cft": 0.028316846592, "cuft": 0.028316846592, "cu.ft": 0.028316846592, "ft3": 0.028316846592,
           "ft³": 0.028316846592}
_LENGTH = {"mm": 1, "cm": 10, "m": 1000, "inch": 25.4, "inches": 25.4, "in": 25.4, '"': 25.4, "인치": 25.4}


def _unit_pattern(units) -> str:
    return "|".join(re.escape(u) for u in sorted(units, key=len, reverse=True))


_WEIGHT_RE = re.compile(rf"({NUM})\s*({_unit_pattern(_WEIGHT)})(?![a-z])", re.I)
_VOLUME_RE = re.compile(rf"({NUM})\s*({_unit_pattern(_VOLUME)})(?![a-z])", re.I)


def parse_weight_kg(text: str) -> Parsed | None:
    text = clean(text)
    match = _WEIGHT_RE.search(text)
    if not match:
        return None
    kg = to_number(match[1]) * _WEIGHT[match[2].lower()]
    return Parsed(round(kg, 3), 1.0, match[0])


def parse_cbm(text: str) -> Parsed | None:
    text = clean(text)
    match = _VOLUME_RE.search(text)
    if not match:
        return None
    return Parsed(round(to_number(match[1]) * _VOLUME[match[2].lower()], 4), 1.0, match[0])


_LEN_UNIT = _unit_pattern(_LENGTH)
_DIMS_RE = re.compile(
    rf"({NUM})\s*({_LEN_UNIT})?\s*[x*]\s*({NUM})\s*({_LEN_UNIT})?\s*[x*]\s*({NUM})\s*({_LEN_UNIT})?(?![a-z])", re.I
)
_LWH_RE = re.compile(rf"L\s*[:=]?\s*({NUM})\s*.*?W\s*[:=]?\s*({NUM})\s*.*?H\s*[:=]?\s*({NUM})\s*({_LEN_UNIT})?", re.I)


def parse_dims_mm(text: str) -> Parsed | None:
    """'310 x 450 x 270mm', '31x45x27 cm', '12 x 18 x 10 inch', 'L310 W450 H270' → (L, W, H) mm 정수."""
    text = clean(text)
    match = _DIMS_RE.search(text)
    if match:
        numbers = [to_number(match[i]) for i in (1, 3, 5)]
        unit = match[6] or match[4] or match[2]
    else:
        match = _LWH_RE.search(text)
        if not match:
            return None
        numbers = [to_number(match[i]) for i in (1, 2, 3)]
        unit = match[4]
    conf = 1.0
    if unit is None:  # 단위 없음 → mm로 보되 확신도를 낮춰 '검토 필요'로 가게 한다
        unit, conf = "mm", 0.6
    factor = _LENGTH[unit.lower()]
    return Parsed(tuple(int(round(n * factor)) for n in numbers), conf, match[0])


# ------------------------------------------------------------------ 수량·포장

_PACK_UNITS = {
    "CTNS": ("ctns", "ctn", "cartons", "carton", "boxes", "box", "bxs", "박스", "상자", "카톤", "개박스"),
    "PLTS": ("plts", "plt", "pallets", "pallet", "팔레트", "파레트", "빠레트"),
    "PKGS": ("pkgs", "pkg", "packages", "package"),
    "CASES": ("cases", "case", "케이스"),
    "DRUMS": ("drums", "drum", "드럼"),
    "CRATES": ("crates", "crate", "크레이트"),
    "PCS": ("pcs", "pc", "pieces", "ea", "개", "피스"),
}
_UNIT_TO_CODE = {u: code for code, units in _PACK_UNITS.items() for u in units}
_QTY_RE = re.compile(rf"({NUM})\s*({_unit_pattern(_UNIT_TO_CODE)})(?![a-z가-힣])", re.I)
PACKING_OF_UNIT = {"CTNS": "Carton", "PLTS": "Pallet", "CRATES": "Crate", "DRUMS": "Drum"}
_PACKING_WORDS = {
    "Carton": ("carton", "ctn", "box", "박스", "상자", "카톤"),
    "Pallet": ("pallet", "plt", "팔레트", "파레트"),
    "Crate": ("crate", "크레이트", "나무상자"),
    "Drum": ("drum", "드럼"),
}


def parse_qty(text: str) -> Parsed | None:
    """'10박스' → (10, 'CTNS'). 포장 단위 수량만 1.0, 낱개(PCS)는 0.6 (6장 qty는 포장 수량 기준)."""
    text = clean(text)
    best = None
    for match in _QTY_RE.finditer(text):
        code = _UNIT_TO_CODE[match[2].lower()]
        conf = 0.6 if code == "PCS" else 1.0
        if best is None or conf > best.conf:
            best = Parsed((int(round(to_number(match[1]))), code), conf, match[0])
    return best


def parse_packing(text: str) -> Parsed | None:
    lowered = clean(text).lower()
    found = [code for code, words in _PACKING_WORDS.items() if any(w in lowered for w in words)]
    return Parsed(found[0], 1.0, text) if len(found) == 1 else None


# ------------------------------------------------------------------ 날짜·금액

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


def parse_date(text: str) -> Parsed | None:
    """ISO·점·슬래시·한글·영문 날짜. 연도가 없으면 None (정밀도 우선 — 추측하지 않는다)."""
    text = clean(text)
    candidates = []
    m = re.search(r"(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})", text)
    if m:
        candidates.append((int(m[1]), int(m[2]), int(m[3]), m[0]))
    m = re.search(r"(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3})[a-z]*\.?,?\s+(20\d{2})", text)
    if m and m[2].lower() in _MONTHS:
        candidates.append((int(m[3]), _MONTHS[m[2].lower()], int(m[1]), m[0]))
    m = re.search(r"([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(20\d{2})", text)
    if m and m[1].lower() in _MONTHS:
        candidates.append((int(m[3]), _MONTHS[m[1].lower()], int(m[2]), m[0]))
    for year, month, day, raw in candidates:
        try:
            return Parsed(date(year, month, day).isoformat(), 1.0, raw)
        except ValueError:
            continue
    return None


_CCY_WORDS = {"USD": ("usd", "us$", "u$", "$", "달러", "미화", "불"), "KRW": ("krw", "₩", "원", "won"),
              "EUR": ("eur", "€", "유로"), "JPY": ("jpy", "¥", "엔"), "CNY": ("cny", "rmb", "위안")}
_CCY_TOKEN = _unit_pattern({w for ws in _CCY_WORDS.values() for w in ws})
_WORD_TO_CCY = {w: c for c, ws in _CCY_WORDS.items() for w in ws}
_MONEY_BEFORE = re.compile(rf"({_CCY_TOKEN})\s*({NUM})", re.I)
_MONEY_AFTER = re.compile(rf"({NUM})\s*(만\s*)?({_CCY_TOKEN})(?![a-z])", re.I)


def parse_money(text: str) -> Parsed | None:
    """'USD 3,000' / '$3,000' / '3,000달러' / '500만원' → (금액, 통화)."""
    text = clean(text)
    m = _MONEY_BEFORE.search(text)
    if m:
        return Parsed((to_number(m[2]), _WORD_TO_CCY[m[1].lower()]), 1.0, m[0])
    m = _MONEY_AFTER.search(text)
    if m:
        amount = to_number(m[1]) * (10_000 if m[2] else 1)
        return Parsed((amount, _WORD_TO_CCY[m[3].lower()]), 1.0, m[0])
    return None


# ------------------------------------------------------------------ 코드값

INCOTERMS = ("FOB", "CIF", "DDP", "EXW")  # 6장 코드값
_ALL_TERMS = ("EXW", "FCA", "FAS", "FOB", "CFR", "CNF", "CIF", "CPT", "CIP", "DAP", "DPU", "DDP")


def find_incoterms(text: str) -> list[str]:
    upper = clean(text).upper()
    return [t for t in _ALL_TERMS if re.search(rf"(?<![A-Z]){t}(?![A-Z])", upper)]


_CONTAINERS = [
    ("40RF", r"40\s*(?:ft|'|피트|’)?\s*(?:rf|rh|reefer|냉동|냉장)"),  # 냉동·냉장 — 일반(GP)보다 먼저 봐야 한다
    ("20RF", r"20\s*(?:ft|'|피트|’)?\s*(?:rf|rh|reefer|냉동|냉장)"),
    ("40HQ", r"40\s*(?:ft|'|피트|’)?\s*(?:hq|hc|high\s*cube|하이\s*큐브)"),
    ("20FT GP", r"20\s*(?:ft|'|피트|’)?\s*(?:gp|dc|dv|dry|일반)?(?:\s*(?:컨테이너|container|cntr))?"),
    ("40FT GP", r"40\s*(?:ft|'|피트|’)?\s*(?:gp|dc|dv|dry|일반)?(?:\s*(?:컨테이너|container|cntr))?"),
]


def parse_container(text: str) -> Parsed | None:
    """'40HQ*1', '40FT HC 1대', "20' GP", 'LCL' → 컨테이너 마스터 코드."""
    cleaned = clean(text)
    upper = cleaned.upper()
    if re.search(r"(?<![A-Z])LCL(?![A-Z])|혼재", upper):
        return Parsed("LCL", 1.0, "LCL")
    for code, pattern in _CONTAINERS:
        m = re.search(rf"(?<!\d){pattern}", cleaned, re.I)
        if m:
            # '20'만 있고 ft/피트/GP 등 표시가 없으면 컨테이너가 아닐 수 있다
            explicit = re.search(r"ft|'|피트|’|gp|dc|dv|dry|hq|hc|cube|rf|rh|reefer|냉동|냉장|컨테이너|container|cntr", m[0], re.I)
            if explicit:
                return Parsed(code, 1.0, m[0])
    return None


_COUNT_AFTER = re.compile(r"^\s*'?\s*(?:x|\*)\s*(\d{1,2})(?!\d)|^[^\d]{0,12}?(\d{1,2})\s*(?:대|units?|cntrs?|containers?|개|ea)(?![a-z])", re.I)
_COUNT_BEFORE = re.compile(r"(\d{1,2})\s*(?:x|\*)\s*$", re.I)


def parse_container_count(text: str) -> int | None:
    """컨테이너 표기 바로 옆의 대수: '40HQ 2대', '40HQ x 2', '2 x 40HQ', "20' GP*3" → 2, 2, 2, 3. 없으면 None."""
    cleaned = clean(text)
    for code, pattern in _CONTAINERS:
        m = re.search(rf"(?<!\d){pattern}", cleaned, re.I)
        if not m:
            continue
        tail = cleaned[m.end():m.end() + 20]
        after = _COUNT_AFTER.search(tail)
        if after:
            return int(after[1] or after[2])
        word = re.match(r"^[^\d가-힣]{0,3}(?:컨테이너\s*)?(한|두|세|네|다섯)\s*대", tail)
        if word:
            return {"한": 1, "두": 2, "세": 3, "네": 4, "다섯": 5}[word[1]]
        before = _COUNT_BEFORE.search(cleaned[max(0, m.start() - 8):m.start()])
        if before:
            return int(before[1])
        return None
    return None


# ------------------------------------------------------------------ 항구 (UN/LOCODE)

LOCODE = re.compile(r"^[A-Z]{2}[A-Z2-9]{3}$")


@lru_cache
def ports() -> dict[str, list[str]]:
    data = json.loads((DATA / "ports.json").read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if not k.startswith("_")}


@lru_cache
def _name_index() -> dict[str, str]:
    return {re.sub(r"\s+", " ", name.upper()): code for code, names in ports().items() for name in names}


def port_code(text: str | None) -> str | None:
    """지명·코드 문자열 → UN/LOCODE. 여러 항구가 섞여 있거나 못 찾으면 None."""
    if not text:
        return None
    upper = re.sub(r"\s+", " ", clean(text).upper())
    token = re.fullmatch(r"\s*([A-Z]{2}\s?[A-Z2-9]{3})\s*", upper)
    if token and token[1].replace(" ", "") in ports():
        return token[1].replace(" ", "")
    index = _name_index()
    found = set()
    for name, code in index.items():
        if len(name) <= 2 and name.isascii():  # 'LA' 같은 짧은 영문은 단어 단위로만
            if re.search(rf"(?<![A-Z]){re.escape(name)}(?![A-Z])", upper):
                found.add(code)
        elif re.search(rf"(?<![A-Z]){re.escape(name)}(?![A-Z])" if name.isascii() else re.escape(name), upper):
            found.add(code)
    return found.pop() if len(found) == 1 else None


def port_candidates(text: str | None, limit: int = 3) -> list[str]:
    """매핑 실패 시 후보 제시 (6장) — 문자열 유사도 기반."""
    if not text:
        return []
    key = re.sub(r"\s+", " ", clean(text).upper()).strip()
    words = [key, *key.split()]
    scored: dict[str, float] = {}
    for name, code in _name_index().items():
        for w in words:
            ratio = difflib.SequenceMatcher(None, w, name).ratio()
            if ratio >= 0.6:
                scored[code] = max(scored.get(code, 0), ratio)
    return [c for c, _ in sorted(scored.items(), key=lambda kv: -kv[1])][:limit]
