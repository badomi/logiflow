"""추출 엔진 공통 자료형."""

import json
from dataclasses import dataclass, field
from functools import lru_cache

from .parsers import DATA

# 6장 입력 필드 (models.INPUT_FIELDS와 같은 순서) + 6장 밖이지만 판단에 쓰는 보조 필드
FIELDS = (
    "commodity", "qty", "qtyUnit", "packing", "boxL", "boxW", "boxH", "totalCbm", "pol", "pod",
    "cargoReadyDate", "incoterms", "containerType", "grossWeightKg", "invoiceValue", "ccy", "paymentTerm",
    "customerName", "contactName", "contactEmail",
)
EXTRA_FIELDS = ("pickupLocation", "containerCount", "carrier")  # 6장 밖 보조값: 판단·견적 계산에만 쓰고 quote_inputs에 저장하지 않는다


@dataclass
class MailText:
    """추출 입력용 메일 1통 (DB Mail·.eml 파싱 결과 어느 쪽에서든 만든다)."""

    subject: str
    from_name: str
    from_email: str
    body: str
    role: str = "ORIGINAL"  # ORIGINAL / REPLY
    attachments: list[tuple[str, str]] = field(default_factory=list)  # (파일명, 추출된 텍스트)


@dataclass
class Source:
    """추출 대상 텍스트 한 덩어리. order가 클수록 최신(회신·첨부가 앞 메일 값을 고친다)."""

    order: int
    label: str  # 근거 표시용: '메일 2 · 본문', '메일 1 · 첨부 packing.pdf'
    kind: str  # subject / body / attachment / signature
    text: str


@dataclass
class Candidate:
    field: str
    value: object
    score: float  # 0~1 (FR-202 유사도 점수)
    evidence: str  # 매칭된 원문 구절
    origin: str  # Source.label
    order: int
    method: str  # rule / llm / signature / header
    unparsed: bool = False  # 값을 찾았지만 형식에 맞게 해석하지 못함 → 저장하지 않고 '확인 필요'로 원문을 보여 준다


@lru_cache
def dictionary() -> dict:
    return json.loads((DATA / "fields.json").read_text(encoding="utf-8"))


def threshold(field_name: str) -> float:
    table = dictionary()["thresholds"]
    return float(table.get(field_name, table["default"]))


UNPARSED_SCORE = 0.5  # 해석 실패 후보의 점수 — 어떤 임계치보다 낮아 저장되지 않고, 제대로 읽힌 후보가 있으면 항상 진다

# 규칙 라벨 그룹 → 그 그룹이 채우는 6장 필드 (해석 실패를 어느 필드에 '확인 필요'로 남길지)
GROUP_FIELDS = {
    "commodity": ("commodity",), "qty": ("qty",), "qtyUnit": ("qty",), "packing": ("packing",),
    "box": ("boxL",), "volume": ("totalCbm",), "totalCbm": ("totalCbm",), "grossWeightKg": ("grossWeightKg",),
    "pol": ("pol",), "pod": ("pod",), "cargoReadyDate": ("cargoReadyDate",), "incoterms": ("incoterms",),
    "containerType": ("containerType",), "invoiceValue": ("invoiceValue",), "ccy": ("ccy",),
    "paymentTerm": ("paymentTerm",), "cargoDetail": (), "pickupLocation": (), "carrier": (),
}
