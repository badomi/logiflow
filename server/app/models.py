"""DB 테이블 정의 (A트랙이 쓰는 저장 인터페이스).

- 이관을 고려해 특정 DB 전용 기능을 쓰지 않는다. Oracle 제약에 맞춰
  이름은 30자 이내, 예약어(references, size 등)는 피하고, 문자열 길이를 모두 지정한다.
- 최종 스키마는 B트랙(DB 담당)과 합의해 확정한다.
- 메일 원문(.eml)과 첨부파일은 파일 저장소에 두고, DB에는 경로만 저장한다.
"""

from datetime import datetime, timezone

from sqlalchemy import Date, DateTime, Float, ForeignKey, Identity, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .db import Base

# 6장 caseStatus 코드값 (FR-107). 상태 전이 관리는 D트랙 담당, A트랙은 생성 시 '수신'만 기록한다.
CASE_STATUSES = ("수신", "파싱중", "정보부족", "보완대기", "재파싱", "계산완료", "승인대기", "발송완료", "보류", "실패")


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class CaseSeq(Base):
    """케이스 ID 날짜별 순번 (FR-102). 날짜마다 한 줄, 발급할 때마다 last_seq를 1 올린다."""

    __tablename__ = "case_seq"

    date_key: Mapped[str] = mapped_column(String(8), primary_key=True)  # YYYYMMDD
    last_seq: Mapped[int] = mapped_column(Integer, default=0)


class Case(Base):
    __tablename__ = "cases"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_id: Mapped[str] = mapped_column(String(20), unique=True)  # LQ-YYYY-MMDD-NNN
    status: Mapped[str] = mapped_column(String(20), default="수신")
    created_by: Mapped[str | None] = mapped_column(String(320))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    mails: Mapped[list["Mail"]] = relationship(back_populates="case", order_by="Mail.id")
    quote_input: Mapped["QuoteInput | None"] = relationship(back_populates="case", uselist=False)
    quotes: Mapped[list["Quote"]] = relationship(back_populates="case", order_by="Quote.version_no")
    extraction_fields: Mapped[list["ExtractionField"]] = relationship(back_populates="case", order_by="ExtractionField.id")
    events: Mapped[list["CaseEvent"]] = relationship(back_populates="case", order_by="CaseEvent.id")


class Mail(Base):
    """케이스에 속한 메일 1통. 최초 요청 메일(ORIGINAL)과 병합된 회신(REPLY)이 모두 여기에 쌓인다."""

    __tablename__ = "mails"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_pk: Mapped[int] = mapped_column(ForeignKey("cases.id"))
    role: Mapped[str] = mapped_column(String(10))  # ORIGINAL / REPLY
    source: Mapped[str] = mapped_column(String(10))  # ADDIN / EML
    # 중복 등록 방지 기준 (FR-105). 같은 메일은 두 번 저장되지 않는다.
    internet_message_id: Mapped[str | None] = mapped_column(String(512), unique=True)
    conversation_id: Mapped[str | None] = mapped_column(String(512), index=True)
    in_reply_to: Mapped[str | None] = mapped_column(String(512))
    ref_headers: Mapped[str | None] = mapped_column(Text)  # References 헤더 (공백 구분)
    subject: Mapped[str | None] = mapped_column(String(1000))
    from_name: Mapped[str | None] = mapped_column(String(320))
    from_email: Mapped[str | None] = mapped_column(String(320))
    to_addrs: Mapped[str | None] = mapped_column(Text)  # JSON 배열
    cc_addrs: Mapped[str | None] = mapped_column(Text)  # JSON 배열
    received_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    body_text: Mapped[str | None] = mapped_column(Text)
    raw_path: Mapped[str | None] = mapped_column(String(500))  # 원문 .eml 저장 경로 (FR-103)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    case: Mapped[Case] = relationship(back_populates="mails")
    attachments: Mapped[list["Attachment"]] = relationship(back_populates="mail", order_by="Attachment.id")


class Attachment(Base):
    __tablename__ = "attachments"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    mail_pk: Mapped[int] = mapped_column(ForeignKey("mails.id"))
    filename: Mapped[str] = mapped_column(String(500))
    content_type: Mapped[str | None] = mapped_column(String(200))
    size_bytes: Mapped[int] = mapped_column(Integer)
    sha256: Mapped[str] = mapped_column(String(64))
    storage_path: Mapped[str] = mapped_column(String(500))

    mail: Mapped[Mail] = relationship(back_populates="attachments")


class CaseEvent(Base):
    """케이스 이력. 누가(actor) 언제 무엇을 했는지 남긴다 (NFR-05 추적성, FR-602 타임라인 재료)."""

    __tablename__ = "case_events"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_pk: Mapped[int] = mapped_column(ForeignKey("cases.id"))
    event_type: Mapped[str] = mapped_column(String(40))
    actor: Mapped[str | None] = mapped_column(String(320))
    detail: Mapped[str | None] = mapped_column(Text)  # JSON
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    case: Mapped[Case] = relationship(back_populates="events")


# =====================================================================================
# 견적 산출 데이터 — 요구사항 정의서 6장 "데이터 요구사항"
#
# 6장 필드는 정의서의 필드명·타입을 그대로 컬럼으로 쓴다 (commodity, qtyUnit, cargoReadyDate …).
# 대소문자가 섞인 이름은 SQLAlchemy가 자동으로 따옴표를 붙여 만들기 때문에 Oracle로 옮겨도 이름이 유지된다.
# 6장 이외의 시스템 컬럼(case_pk 등)은 기존 A트랙 테이블 규칙(snake_case)을 따른다.
#
# 6장 "필수"는 DB 제약(NOT NULL)으로 걸지 않는다. 케이스는 메일 선택 즉시 만들어지고 값은 추출·보완 회신으로
# 나중에 채워지기 때문이다. 필수·형식 판정은 validation.py(룰 베이스)가 하고, 빠진 항목은 보완 요청 문항이 된다.
# caseStatus는 기존 cases.status 컬럼이 같은 역할을 한다 (값 목록은 6장과 동일).
# =====================================================================================


class QuoteInput(Base):
    """케이스 1건의 6장 필드 값 (케이스당 1행). 추출(LLM)이 채우고 검증(룰)이 판정한다."""

    __tablename__ = "quote_inputs"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_pk: Mapped[int] = mapped_column(ForeignKey("cases.id"), unique=True)

    # --- 견적 산출 최소 필드 (6장 표 1) ---
    commodity: Mapped[str | None] = mapped_column(String(500))
    qty: Mapped[int | None] = mapped_column(Integer)
    qtyUnit: Mapped[str | None] = mapped_column(String(20))
    packing: Mapped[str | None] = mapped_column(String(20))  # Carton / Pallet / Crate / Drum
    boxL: Mapped[int | None] = mapped_column(Integer)  # mm
    boxW: Mapped[int | None] = mapped_column(Integer)  # mm
    boxH: Mapped[int | None] = mapped_column(Integer)  # mm
    totalCbm: Mapped[float | None] = mapped_column(Float)
    pol: Mapped[str | None] = mapped_column(String(5))  # UN/LOCODE
    pod: Mapped[str | None] = mapped_column(String(5))  # UN/LOCODE
    cargoReadyDate: Mapped[str | None] = mapped_column(String(10))  # 날짜(ISO) YYYY-MM-DD
    incoterms: Mapped[str | None] = mapped_column(String(3))  # FOB / CIF / DDP / EXW
    containerType: Mapped[str | None] = mapped_column(String(20))  # container_types.code 값만
    grossWeightKg: Mapped[float | None] = mapped_column(Float)
    invoiceValue: Mapped[float | None] = mapped_column(Float)
    ccy: Mapped[str | None] = mapped_column(String(3))
    paymentTerm: Mapped[str | None] = mapped_column(String(200))
    customerName: Mapped[str | None] = mapped_column(String(300))
    contactName: Mapped[str | None] = mapped_column(String(200))
    contactEmail: Mapped[str | None] = mapped_column(String(320))

    # --- 견적 산출·표기용 파생 필드 (6장 표 2, 요율/룰 마스터에서 결정) ---
    quoteCurrency: Mapped[str | None] = mapped_column(String(3))
    exchangeRateAsOf: Mapped[str | None] = mapped_column(String(10))
    transitTime: Mapped[int | None] = mapped_column(Integer)
    freeTimeDet: Mapped[int | None] = mapped_column(Integer)
    freeTimeDem: Mapped[int | None] = mapped_column(Integer)
    validityDays: Mapped[int | None] = mapped_column(Integer)

    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, onupdate=utcnow)

    case: Mapped[Case] = relationship(back_populates="quote_input")


# 6장 입력 필드 목록 (추출·검증·API가 같은 순서로 쓴다)
INPUT_FIELDS = (
    "commodity", "qty", "qtyUnit", "packing", "boxL", "boxW", "boxH", "totalCbm", "pol", "pod",
    "cargoReadyDate", "incoterms", "containerType", "grossWeightKg", "invoiceValue", "ccy", "paymentTerm",
    "customerName", "contactName", "contactEmail",
)
DERIVED_FIELDS = ("quoteCurrency", "exchangeRateAsOf", "transitTime", "freeTimeDet", "freeTimeDem", "validityDays")


class ExtractionField(Base):
    """필드별 추출 근거 (FR-202: 값·매칭된 원문 구절·점수) + 담당자 수정 이력 (FR-206).

    추출할 때마다 run_no가 올라가며 21개 필드 전부 한 줄씩 남는다 (NFR-05 판단 근거 로그).
    method='manual'은 담당자가 고친 값 — 이후 재추출이 그 필드를 덮어쓰지 않는다.
    """

    __tablename__ = "extraction_fields"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_pk: Mapped[int] = mapped_column(ForeignKey("cases.id"), index=True)
    run_no: Mapped[int] = mapped_column(Integer)
    field_name: Mapped[str] = mapped_column(String(30))  # 6장 필드명
    value: Mapped[str | None] = mapped_column(Text)  # JSON 문자열 (숫자·문자 구분 유지)
    score: Mapped[float] = mapped_column(Float, default=0.0)  # 0~1
    evidence: Mapped[str | None] = mapped_column(Text)  # 매칭된 원문 구절
    origin: Mapped[str | None] = mapped_column(String(200))  # '메일 2 · 회신 본문' 등
    method: Mapped[str | None] = mapped_column(String(20))  # rule / llm / rule+llm / signature / header / manual
    status: Mapped[str] = mapped_column(String(10))  # filled / review / missing
    actor: Mapped[str | None] = mapped_column(String(320))  # manual일 때 수정한 담당자
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    case: Mapped[Case] = relationship(back_populates="extraction_fields")


class ContainerType(Base):
    """컨테이너 마스터. containerType은 여기 등록된 code만 허용한다 (6장 제약). 'LCL'은 혼재 화물 코드."""

    __tablename__ = "container_types"

    code: Mapped[str] = mapped_column(String(20), primary_key=True)
    label: Mapped[str | None] = mapped_column(String(100))
    max_gross_kg: Mapped[float | None] = mapped_column(Float)  # 허용중량 초과 검사 기준
    max_cbm: Mapped[float | None] = mapped_column(Float)


class Rate(Base):
    """요율/부대비용 마스터 1행 = 견적서의 비용 1줄 (OCEAN FREIGHT, THC, DOCUMENT FEE …).

    pol·pod·containerType·incoterms가 비어 있으면 '모든 값에 적용'으로 본다.
    basis: PER_RT(LCL R/T) / PER_CNTR / PER_BL / PER_TRIP / AT_COST(실비, 금액 없이 산정식만 표기)
    """

    __tablename__ = "rates"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    charge_code: Mapped[str] = mapped_column(String(40))
    charge_label: Mapped[str] = mapped_column(String(100))  # 견적서에 찍히는 이름
    pol: Mapped[str | None] = mapped_column(String(5))
    pod: Mapped[str | None] = mapped_column(String(5))
    container_type: Mapped[str | None] = mapped_column(String(20))
    incoterms: Mapped[str | None] = mapped_column(String(3))
    basis: Mapped[str] = mapped_column(String(10))
    unit_price: Mapped[float | None] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3), default="KRW")
    remark: Mapped[str | None] = mapped_column(String(300))
    sort_no: Mapped[int] = mapped_column(Integer, default=100)
    valid_from: Mapped[datetime | None] = mapped_column(Date)
    valid_until: Mapped[datetime | None] = mapped_column(Date)
    source: Mapped[str | None] = mapped_column(String(200))  # 요율 출처 (어느 요율표인지)


class LaneRule(Base):
    """구간별 파생 필드 룰 (6장 표 2: transitTime·freeTimeDet/Dem·validityDays·quoteCurrency)"""

    __tablename__ = "lane_rules"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    pol: Mapped[str | None] = mapped_column(String(5))
    pod: Mapped[str | None] = mapped_column(String(5))
    container_type: Mapped[str | None] = mapped_column(String(20))
    quote_currency: Mapped[str] = mapped_column(String(3), default="USD")
    transit_days: Mapped[int | None] = mapped_column(Integer)
    free_time_det: Mapped[int | None] = mapped_column(Integer)
    free_time_dem: Mapped[int | None] = mapped_column(Integer)
    validity_days: Mapped[int] = mapped_column(Integer, default=14)


class Quote(Base):
    """견적서 1버전. 재계산할 때마다 버전이 올라가고 이전 버전은 남는다."""

    __tablename__ = "quotes"
    __table_args__ = (UniqueConstraint("case_pk", "version_no"),)

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    case_pk: Mapped[int] = mapped_column(ForeignKey("cases.id"))
    version_no: Mapped[int] = mapped_column(Integer)
    valid_until: Mapped[str | None] = mapped_column(String(10))
    xlsx_path: Mapped[str | None] = mapped_column(String(500))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    case: Mapped[Case] = relationship(back_populates="quotes")
    items: Mapped[list["QuoteItem"]] = relationship(back_populates="quote", order_by="QuoteItem.id")
    totals: Mapped[list["QuoteTotal"]] = relationship(back_populates="quote", order_by="QuoteTotal.id")


class QuoteItem(Base):
    """견적서 비용 1줄: 단가(rate) × 수량(qty) = 금액(amount). 실비(AT COST)면 금액 없이 비고만."""

    __tablename__ = "quote_items"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    quote_pk: Mapped[int] = mapped_column(ForeignKey("quotes.id"))
    charge_code: Mapped[str] = mapped_column(String(40))
    charge_label: Mapped[str] = mapped_column(String(100))
    basis: Mapped[str] = mapped_column(String(10))
    rate: Mapped[float | None] = mapped_column(Float)
    qty: Mapped[float | None] = mapped_column(Float)
    amount: Mapped[float | None] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3))
    remark: Mapped[str | None] = mapped_column(String(300))

    quote: Mapped[Quote] = relationship(back_populates="items")


class QuoteTotal(Base):
    """통화별 합계. 양식의 'TOTAL (USD+KRW)'처럼 한 견적서에 통화가 섞인다."""

    __tablename__ = "quote_totals"

    id: Mapped[int] = mapped_column(Integer, Identity(), primary_key=True)
    quote_pk: Mapped[int] = mapped_column(ForeignKey("quotes.id"))
    currency: Mapped[str] = mapped_column(String(3))
    total: Mapped[float] = mapped_column(Float)

    quote: Mapped[Quote] = relationship(back_populates="totals")
