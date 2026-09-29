"""추출: 케이스의 메일(원문 + 병합된 회신)을 로컬 LLM으로 읽어 6장 필드를 채운다.

pipeline.py의 Extractor 규격 구현. 흐름:
  1) 메일 본문 정리 — 인용문(이전 메일) 제거, Outlook 번호 목록 갈라짐 복구 (docs/트랙간_전달사항.md 2-1)
  2) 로컬 LLM에 "적힌 값만, 없으면 null" 조건으로 JSON 추출 요청 (추측 금지)
  3) 룰 베이스 정규화 — 형식 통일(UN/LOCODE, ISO 날짜, 코드값), 판단이 필요한 경우는 비워 두고 사유 기록
  4) quote_inputs에 저장. 새로 읽은 값이 있으면 갱신, 없으면 기존 값 유지 (회신으로 부족분 병합, MUST-SHIP ③)

돌려주는 dict는 case_events(PIPELINE_RUN) 상세에 그대로 남는다 → 판단 근거 로그 (NFR-05).
"""

import difflib
import re
from datetime import date

from sqlalchemy.orm import Session

from .llm import LlmClient, OllamaClient
from .models import INPUT_FIELDS, Case, QuoteInput

MAX_BODY_CHARS = 12_000  # 60초 안에 끝나도록 입력 길이 제한

# ------------------------------------------------------------------ 메일 본문 정리

_QUOTE_STARTS = (
    re.compile(r"^_{8,}\s*$"),  # Outlook 웹 회신 구분선
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}", re.I),
    re.compile(r"^-{2,}\s*원본 메시지\s*-{2,}"),
    re.compile(r"^(보낸 사람|From)\s*:", re.I),
    re.compile(r".*님이 작성:\s*$"),  # Gmail 한국어
    re.compile(r"^On .+wrote:\s*$"),  # Gmail 영어
)


def strip_quoted(body: str) -> str:
    """회신 본문에서 인용된 이전 메일을 잘라낸다 (첫 인용 표시 이후 전부 + '>' 줄).

    서명은 인용 표시보다 위에 있으므로 남는다 (FR-208 연락처 추출용).
    """
    kept: list[str] = []
    for line in body.replace("\r\n", "\n").split("\n"):
        stripped = line.strip()
        if any(p.match(stripped) for p in _QUOTE_STARTS):
            break
        if stripped.startswith(">"):
            continue
        kept.append(line)
    return "\n".join(kept).strip()


_LONE_NUMBER = re.compile(r"^\s*(\d{1,2}[.)]|[-•·*])\s*$")


def join_broken_list(body: str) -> str:
    """Outlook HTML 목록 변환으로 `1.` 과 내용이 다른 줄로 갈라진 것을 한 줄로 합친다."""
    lines = body.split("\n")
    out: list[str] = []
    i = 0
    while i < len(lines):
        if _LONE_NUMBER.match(lines[i]) and i + 1 < len(lines) and lines[i + 1].strip():
            out.append(f"{lines[i].strip()} {lines[i + 1].strip()}")
            i += 2
            continue
        out.append(lines[i])
        i += 1
    return "\n".join(out)


def mail_texts(case: Case) -> str:
    """케이스의 모든 메일을 시간 순서로 이어 붙인 LLM 입력."""
    parts = []
    for index, mail in enumerate(case.mails, start=1):
        label = "최초 요청" if mail.role == "ORIGINAL" else "회신"
        body = join_broken_list(strip_quoted(mail.body_text or ""))
        parts.append(
            f"[메일 {index} · {label}]\n보낸 사람: {mail.from_name or ''} <{mail.from_email or ''}>\n"
            f"제목: {mail.subject or ''}\n본문:\n{body}"
        )
    return "\n\n".join(parts)[:MAX_BODY_CHARS]


# ------------------------------------------------------------------ LLM 요청

SYSTEM_PROMPT = """너는 해운 포워딩 회사의 견적 요청 메일에서 값을 뽑는 도구다.
규칙:
- 메일에 실제로 적힌 값만 뽑는다. 적혀 있지 않으면 반드시 null. 추측·계산·일반 상식으로 채우지 않는다.
- 메일이 여러 통이면 뒤의 메일(회신)이 앞 메일의 값을 고치거나 보충한다. 가장 최신 값을 쓴다.
- pol·pod는 UN/LOCODE 5자리를 확실히 알 때만 코드로, 아니면 null. 적힌 항구·지명은 polName·podName에 원문 그대로.
- 선적항이 정해지지 않았거나 "알아서 정해 달라"는 경우 pol은 null.
- cargoReadyDate는 YYYY-MM-DD. 연도가 없거나 "다음 주" 같은 상대 표현이면 null.
- incoterms는 적힌 조건 문자열 그대로(예: "FOB 및 CIF"). containerType도 적힌 그대로(예: "40FT HC", "LCL").
- 숫자 필드는 숫자만(쉼표·단위 제외). grossWeightKg는 kg, 박스 규격 boxL/W/H는 mm.
- 품목이 둘 이상이고 품목마다 수량·규격이 따로 적혀 있으면 multipleItems=true.
- evidence에는 값을 뽑은 근거 문장을 원문 그대로 적는다(필드명: 문장).
"""

_NUM = {"type": ["number", "null"]}
_INT = {"type": ["integer", "null"]}
_STR = {"type": ["string", "null"]}

EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "commodity": _STR, "qty": _INT, "qtyUnit": _STR, "packing": _STR,
        "boxL": _INT, "boxW": _INT, "boxH": _INT, "totalCbm": _NUM,
        "pol": _STR, "polName": _STR, "pod": _STR, "podName": _STR, "pickupLocation": _STR,
        "cargoReadyDate": _STR, "incoterms": _STR, "containerType": _STR,
        "grossWeightKg": _NUM, "invoiceValue": _NUM, "ccy": _STR, "paymentTerm": _STR,
        "customerName": _STR, "contactName": _STR, "contactEmail": _STR,
        "multipleItems": {"type": "boolean"},
        "evidence": {"type": "object", "additionalProperties": {"type": "string"}},
    },
    "required": ["commodity", "qty", "pol", "pod", "incoterms", "multipleItems", "evidence"],
}

# ------------------------------------------------------------------ 정규화 (룰 베이스)

# 자주 나오는 항구 → UN/LOCODE. LLM이 코드를 못 줬을 때 지명으로 찾는다. (사내 마스터로 교체 예정)
PORTS = {
    "KRPUS": ("부산", "BUSAN", "PUSAN", "부산항", "부산북항", "부산신항"),
    "KRINC": ("인천", "INCHEON", "인천항"),
    "KRKAN": ("광양", "GWANGYANG", "KWANGYANG"),
    "KRPTK": ("평택", "PYEONGTAEK"),
    "AUSYD": ("시드니", "SYDNEY"),
    "AUMEL": ("멜버른", "MELBOURNE"),
    "CNSHA": ("상하이", "SHANGHAI"),
    "CNNGB": ("닝보", "NINGBO"),
    "HKHKG": ("홍콩", "HONG KONG"),
    "TWKEL": ("지룽", "기륭", "KEELUNG"),
    "SGSIN": ("싱가포르", "SINGAPORE"),
    "JPTYO": ("도쿄", "TOKYO"),
    "USLAX": ("로스앤젤레스", "LOS ANGELES", "LA"),
    "USLGB": ("롱비치", "LONG BEACH"),
    "USNYC": ("뉴욕", "NEW YORK"),
    "ITGOA": ("제노바", "GENOA"),
    "NLRTM": ("로테르담", "ROTTERDAM"),
    "DEHAM": ("함부르크", "HAMBURG"),
    "AEJEA": ("제벨알리", "JEBEL ALI"),
    "VNSGN": ("호치민", "HO CHI MINH"),
}
_NAME_TO_CODE = {name.upper(): code for code, names in PORTS.items() for name in names}
LOCODE = re.compile(r"^[A-Z]{2}[A-Z2-9]{3}$")


def port_code(code: str | None, name: str | None) -> str | None:
    if code and LOCODE.match(code.strip().upper()) and code.strip().upper() in PORTS:
        return code.strip().upper()
    if code and LOCODE.match(code.strip().upper()):
        return code.strip().upper()  # 목록에 없는 코드라도 형식이 맞으면 받는다 (검증에서 확인)
    if name:
        key = re.sub(r"[,()]|\bPORT\b|\bKOREA\b|\bAUSTRALIA\b|항$", " ", name.upper()).strip()
        for token in [key, *key.split()]:
            if token in _NAME_TO_CODE:
                return _NAME_TO_CODE[token]
    return None


def port_candidates(name: str | None, limit: int = 3) -> list[str]:
    """매핑 실패 시 후보 제시용 (6장 pol·pod 제약) — 문자열 유사도 기반 (FR-202 방식)."""
    if not name:
        return []
    matches = difflib.get_close_matches(name.upper(), list(_NAME_TO_CODE), n=limit * 2, cutoff=0.5)
    seen: list[str] = []
    for m in matches:
        code = _NAME_TO_CODE[m]
        if code not in seen:
            seen.append(code)
    return seen[:limit]


_PACKING = {
    "Carton": ("CARTON", "CTN", "CTNS", "BOX", "BOXES", "박스", "상자", "카톤"),
    "Pallet": ("PALLET", "PLT", "PLTS", "팔레트", "파레트"),
    "Crate": ("CRATE", "크레이트", "나무상자"),
    "Drum": ("DRUM", "드럼"),
}
_QTY_UNIT = {"CTNS": ("CTN", "CTNS", "CARTON", "BOX", "BOXES", "박스", "상자"), "PLTS": ("PLT", "PLTS", "PALLET", "팔레트")}
_CONTAINER = {
    "LCL": ("LCL", "혼재"),
    "20FT GP": ("20GP", "20FTGP", "20FT", "20DC", "20DRY", "20피트"),
    "40FT GP": ("40GP", "40FTGP", "40FT", "40DC", "40DRY"),
    "40HQ": ("40HQ", "40HC", "40FTHC", "40FTHQ", "40HIGHCUBE", "40하이큐브"),
}
_CCY = {"USD": ("USD", "US$", "$", "달러", "미화"), "KRW": ("KRW", "원", "₩", "WON"), "EUR": ("EUR", "€", "유로")}
INCOTERMS = ("FOB", "CIF", "DDP", "EXW")


def _lookup(value: str | None, table: dict) -> str | None:
    if not value:
        return None
    key = re.sub(r"[\s'’\"\-.]", "", value).upper()
    for code, names in table.items():
        if key in {re.sub(r"[\s'’\"\-.]", "", n).upper() for n in names}:
            return code
    return None


def _to_float(value) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d[\d,]*\.?\d*", str(value))
    return float(match.group().replace(",", "")) if match else None


def _to_int(value) -> int | None:
    number = _to_float(value)
    return int(round(number)) if number is not None else None


def _to_date(value) -> str | None:
    if not value:
        return None
    match = re.search(r"(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})", str(value))
    if not match:
        return None
    try:
        return date(int(match[1]), int(match[2]), int(match[3])).isoformat()
    except ValueError:
        return None


def _clean(value) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def normalize(raw: dict) -> tuple[dict, list[dict]]:
    """LLM 원시 값 → 6장 형식. 판단이 필요한 값은 None으로 두고 notes에 사유를 남긴다."""
    notes: list[dict] = []
    out: dict = {k: None for k in INPUT_FIELDS}

    out["commodity"] = _clean(raw.get("commodity"))
    out["qty"] = _to_int(raw.get("qty"))
    unit = _clean(raw.get("qtyUnit"))
    out["qtyUnit"] = _lookup(unit, _QTY_UNIT) or (unit.upper() if unit else None)
    out["packing"] = _lookup(_clean(raw.get("packing")), _PACKING) or _lookup(unit, _PACKING)
    for key in ("boxL", "boxW", "boxH"):
        out[key] = _to_int(raw.get(key))
    out["totalCbm"] = _to_float(raw.get("totalCbm"))

    for side in ("pol", "pod"):
        name = _clean(raw.get(f"{side}Name"))
        out[side] = port_code(_clean(raw.get(side)), name)
        if out[side] is None and name:
            notes.append({"field": side, "code": "PORT_UNMAPPED", "value": name, "candidates": port_candidates(name)})

    out["cargoReadyDate"] = _to_date(raw.get("cargoReadyDate"))
    if raw.get("cargoReadyDate") and not out["cargoReadyDate"]:
        notes.append({"field": "cargoReadyDate", "code": "DATE_UNPARSED", "value": raw.get("cargoReadyDate")})

    terms_raw = (_clean(raw.get("incoterms")) or "").upper()
    found = [t for t in INCOTERMS if re.search(rf"\b{t}\b", terms_raw)]
    if len(found) == 1:
        out["incoterms"] = found[0]
    elif len(found) > 1:
        notes.append({"field": "incoterms", "code": "MULTIPLE_INCOTERMS", "value": found})
    elif terms_raw:
        notes.append({"field": "incoterms", "code": "UNSUPPORTED_INCOTERMS", "value": terms_raw})

    cntr = _clean(raw.get("containerType"))
    out["containerType"] = _lookup(cntr, _CONTAINER) or (cntr.upper() if cntr else None)

    out["grossWeightKg"] = _to_float(raw.get("grossWeightKg"))
    out["invoiceValue"] = _to_float(raw.get("invoiceValue"))
    ccy = _clean(raw.get("ccy"))
    out["ccy"] = _lookup(ccy, _CCY) or (ccy.upper() if ccy and len(ccy) == 3 else None)
    out["paymentTerm"] = _clean(raw.get("paymentTerm"))
    out["customerName"] = _clean(raw.get("customerName"))
    out["contactName"] = _clean(raw.get("contactName"))
    email = _clean(raw.get("contactEmail"))
    out["contactEmail"] = email if email and re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) else None
    return out, notes


# ------------------------------------------------------------------ Extractor 구현


class LlmExtractor:
    """pipeline.Extractor 구현 — register(extractor=LlmExtractor())로 연결."""

    name = "local-llm"

    def __init__(self, client: LlmClient | None = None):
        self._client = client

    @property
    def client(self) -> LlmClient:
        if self._client is None:
            self._client = OllamaClient()
        return self._client

    def extract(self, session: Session, case: Case) -> dict:
        raw = self.client.complete_json(SYSTEM_PROMPT, mail_texts(case), EXTRACTION_SCHEMA)
        values, notes = normalize(raw)

        # 연락처는 서명에서 못 찾으면 최초 요청 메일의 보낸 사람으로 (FR-208: 헤더/서명에서 추출)
        original = case.mails[0] if case.mails else None
        if original is not None:
            values["contactEmail"] = values["contactEmail"] or original.from_email
            values["contactName"] = values["contactName"] or original.from_name or None

        record = case.quote_input or QuoteInput(case=case)
        changed = []
        for key, value in values.items():
            if value is not None and getattr(record, key) != value:
                setattr(record, key, value)
                changed.append(key)
        session.add(record)
        session.flush()

        return {
            "fields": {k: getattr(record, k) for k in INPUT_FIELDS},
            "changed": changed,
            "multipleItems": bool(raw.get("multipleItems")),
            "pickupLocation": _clean(raw.get("pickupLocation")),
            "notes": notes,
            "evidence": raw.get("evidence") or {},
        }
