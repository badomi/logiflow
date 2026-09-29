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
EXTRA_FIELDS = ("pickupLocation", "containerCount")  # 6장 밖 보조값: 판단·견적 계산에만 쓰고 quote_inputs에 저장하지 않는다


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


@lru_cache
def dictionary() -> dict:
    return json.loads((DATA / "fields.json").read_text(encoding="utf-8"))


def threshold(field_name: str) -> float:
    table = dictionary()["thresholds"]
    return float(table.get(field_name, table["default"]))
