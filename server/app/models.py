"""DB 테이블 정의 (A트랙이 쓰는 저장 인터페이스).

- 이관을 고려해 특정 DB 전용 기능을 쓰지 않는다. Oracle 제약에 맞춰
  이름은 30자 이내, 예약어(references, size 등)는 피하고, 문자열 길이를 모두 지정한다.
- 최종 스키마는 B트랙(DB 담당)과 합의해 확정한다.
- 메일 원문(.eml)과 첨부파일은 파일 저장소에 두고, DB에는 경로만 저장한다.
"""

from datetime import datetime, timezone

from sqlalchemy import DateTime, ForeignKey, Identity, Integer, String, Text
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
