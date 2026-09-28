"""API로 주고받는 JSON 형식.

MailSnapshot은 A트랙 "메일 표준 JSON"(v1)이다. 애드인 패널, .eml 리더, 백엔드가 모두 이 모양을 쓴다.
JSON 키는 camelCase(6장 필드 표기와 동일), 파이썬 코드 안에서는 snake_case로 쓴다.
"""

from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field
from pydantic.alias_generators import to_camel


class CamelModel(BaseModel):
    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)


class Address(CamelModel):
    name: str = ""
    email: str = ""


class AttachmentInfo(CamelModel):
    filename: str
    content_type: str | None = None
    size: int
    sha256: str


class MailSnapshot(CamelModel):
    subject: str = ""
    from_: Address = Field(default_factory=Address, alias="from")
    to: list[Address] = []
    cc: list[Address] = []
    received_at: datetime | None = None
    body_text: str = ""
    conversation_id: str | None = None  # Outlook 대화 ID (.eml에는 없음)
    internet_message_id: str | None = None  # Message-ID 헤더 — 중복 방지(FR-105)
    in_reply_to: str | None = None  # 회신 식별(FR-104)
    references: list[str] = []  # 회신 식별(FR-104)
    attachments: list[AttachmentInfo] = []


class MailInput(CamelModel):
    """애드인이 보내는 요청. 원문 .eml(Base64)을 보내면 서버가 파싱하고, snapshot 값으로 보완한다."""

    eml_base64: str | None = None
    snapshot: MailSnapshot | None = None
    actor: str | None = None  # 버튼을 누른 담당자 (이력 기록용)


class EventOut(CamelModel):
    event_type: str
    actor: str | None
    detail: dict | None
    created_at: datetime


class MailOut(CamelModel):
    role: str
    source: str
    snapshot: MailSnapshot
    has_raw: bool


class CaseSummary(CamelModel):
    case_id: str
    status: str
    subject: str
    from_email: str
    created_at: datetime


class CaseDetail(CaseSummary):
    mails: list[MailOut]
    events: list[EventOut]


class CreateCaseResult(CamelModel):
    duplicate: bool  # 이미 등록된 메일이면 True (새 케이스를 만들지 않음)
    case: CaseDetail


class ReplyCandidate(CamelModel):
    case_id: str
    subject: str
    score: float  # 0~1, 높을수록 확실
    reasons: list[str]  # SUBJECT_CASE_ID / HEADER / CONVERSATION


class DraftAttachment(CamelModel):
    filename: str
    url: str  # 패널이 내려받을 백엔드 경로
    size: int


class DraftOut(CamelModel):
    """메일 초안 내용. 패널이 이 값으로 Outlook 작성 창을 띄우고, 발송은 담당자가 한다."""

    kind: str  # SUPPLEMENT(보완 요청) / QUOTE(견적서 송부)
    to: list[Address]
    subject: str
    html_body: str
    attachments: list[DraftAttachment]


class SupplementRequest(CamelModel):
    questions: list[str]  # C트랙 검증 결과의 보완 문항 (연결 전에는 담당자 입력)
    actor: str | None = None


class ActorRequest(CamelModel):
    actor: str | None = None


class MailEventRequest(CamelModel):
    """Outlook 작성 창·보내기 이벤트가 알려 주는 기록 (FR-505 발송 시각·수행자)."""

    event_type: str  # DRAFT_SAVED(초안함 저장) / MAIL_SENT(담당자가 보내기를 누름)
    kind: str | None = None  # SUPPLEMENT / QUOTE (제목으로 판별 못 하면 비움)
    subject: str = ""
    actor: str | None = None
    internet_message_id: str | None = None  # 보낸 메일의 ID — 같은 발송을 두 번 기록하지 않는다
    occurred_at: datetime | None = None  # 실제 발송 시각 (보낸 편지함 메일에서 기록할 때)


class ReplyMatchResult(CamelModel):
    already_registered_case_id: str | None  # 이 메일이 이미 어떤 케이스에 들어가 있으면 그 ID
    candidates: list[ReplyCandidate]
