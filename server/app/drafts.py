"""메일 초안 내용 만들기 — 보완 요청(FR-304·305·510), 견적서 송부(FR-505).

A트랙은 "받는 사람·제목·본문·첨부"를 만들어 패널에 넘기고, 패널이 Outlook 작성 창을 띄운다.
발송은 항상 담당자가 직접 누른다. 이 모듈에는 메일을 보내는 코드가 없다 (정의서 12장 설계 고정).
"""

import json
from datetime import datetime, timedelta, timezone
from html import escape
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from . import storage
from .config import settings
from .models import Case, CaseEvent
from .schemas import Address, DraftAttachment, DraftOut

SIGNATURE_KO = "레오나해운항공㈜ 견적 담당"
SIGNATURE_EN = "LEONA SEA &amp; AIR CO., LTD. Quotation Team"

QUOTE_EXTENSIONS = {".pdf", ".xlsx"}


class DraftError(ValueError):
    """초안을 만들 수 없음 (문항·견적서 없음 등)"""


def _original(case: Case):
    if not case.mails:
        raise DraftError("케이스에 메일이 없습니다.")
    return case.mails[0]


def _recipient(case: Case) -> list[Address]:
    """받는 사람: B트랙이 서명·헤더에서 뽑은 회신 주소(FR-208)가 있으면 그것, 없으면 최초 요청 메일의 발신자."""
    original = _original(case)
    qi = case.quote_input
    if qi is not None and qi.contactEmail:
        return [Address(name=qi.contactName or original.from_name or "", email=qi.contactEmail)]
    return [Address(name=original.from_name or "", email=original.from_email or "")]


def validation_questions(case: Case) -> list[str]:
    """가장 최근 검증(C트랙) 결과의 보완 문항 (FR-304). 검증이 안 돌았으면 빈 목록."""
    for event in reversed(case.events):
        if event.event_type == "PIPELINE_RUN":
            validation = json.loads(event.detail or "{}").get("validation") or {}
            return [q for q in validation.get("questions") or [] if q]
    return []


def _record(session: Session, case: Case, kind: str, actor: str | None, detail: dict) -> None:
    session.add(
        CaseEvent(
            case=case,
            event_type="DRAFT_PREPARED",
            actor=actor,
            detail=json.dumps({"kind": kind, **detail}, ensure_ascii=False),
        )
    )
    session.commit()


def supplement_draft(session: Session, case: Case, questions: list[str], actor: str | None) -> DraftOut:
    """보완 요청 초안. 문항은 누락 항목 그대로 번호를 붙인다 — 임의로 더하거나 빼지 않는다 (FR-305)."""
    questions = [q.strip() for q in questions if q and q.strip()]
    if not questions:  # 패널에서 문항을 비워 보냈으면 검증 결과 문항을 그대로 쓴다 (FR-305)
        questions = validation_questions(case)
    if not questions:
        raise DraftError(
            "보완 요청할 문항이 없어 초안을 만들지 않았습니다. "
            "검증 결과 빠진 항목이 없거나 아직 추출이 끝나지 않았습니다. "
            "직접 물어볼 내용이 있으면 '보완 요청 문항' 칸에 한 줄에 하나씩 입력한 뒤 다시 누르세요."
        )

    subject = f"[{case.case_id}] 견적 보완 요청 / Request for additional information - {_original(case).subject or ''}".strip()
    items = "".join(f"<li>{escape(q)}</li>" for q in questions)
    body = (
        "<p>안녕하세요, 레오나해운항공㈜입니다.<br>"
        "요청하신 견적을 산출하기 위해 아래 항목을 추가로 알려 주시기 바랍니다.</p>"
        f"<ol>{items}</ol>"
        f"<p>회신 시 제목의 케이스 번호({escape(case.case_id)})를 유지해 주시면 빠르게 처리됩니다.<br>"
        f"감사합니다.<br>{SIGNATURE_KO}</p>"
        "<hr>"
        "<p>Hello, this is LEONA SEA &amp; AIR CO., LTD.<br>"
        "To prepare your quotation, please provide the information listed above.<br>"
        f"Please keep the case number ({escape(case.case_id)}) in the subject when replying.<br>"
        f"Thank you.<br>{SIGNATURE_EN}</p>"
    )
    _record(session, case, "SUPPLEMENT", actor, {"subject": subject, "questionCount": len(questions)})
    return DraftOut(kind="SUPPLEMENT", case_id=case.case_id, to=_recipient(case), subject=subject, html_body=body, attachments=[])


MAIL_EVENT_TYPES = {"DRAFT_SAVED", "MAIL_SENT"}

# 담당자가 메일을 보낸 뒤의 케이스 상태 (6장 caseStatus). 상태 전이 규칙의 최종 결정은 D트랙.
STATUS_AFTER_SENT = {"SUPPLEMENT": "보완대기", "QUOTE": "발송완료"}


def record_mail_event(
    session: Session,
    case: Case,
    event_type: str,
    kind: str | None,
    subject: str,
    actor: str | None,
    internet_message_id: str | None = None,
    occurred_at: datetime | None = None,
) -> bool:
    """초안 저장·발송을 이력에 남긴다 (FR-505). 새로 기록했으면 True, 이미 기록된 발송이면 False.

    발송 시각은 보낸 메일의 실제 시각(occurred_at)이 있으면 그 값, 없으면 기록 시각이다.
    이 함수는 기록만 한다. 메일을 보내는 것은 Outlook에서 담당자가 직접 누른 보내기 버튼이다.
    """
    if event_type not in MAIL_EVENT_TYPES:
        raise DraftError(f"알 수 없는 이벤트입니다: {event_type}")
    if kind is None:
        kind = "QUOTE" if "견적서 송부" in subject else "SUPPLEMENT" if "보완 요청" in subject else None
    if internet_message_id and any(
        e.event_type == event_type and (json.loads(e.detail or "{}").get("messageId") == internet_message_id)
        for e in case.events
    ):
        return False

    detail = {"kind": kind, "subject": subject}
    if internet_message_id:
        detail["messageId"] = internet_message_id
    if occurred_at:
        detail["sentAt"] = occurred_at.astimezone(timezone.utc).isoformat()
    if event_type == "MAIL_SENT" and kind in STATUS_AFTER_SENT:
        detail["statusFrom"] = case.status
        case.status = STATUS_AFTER_SENT[kind]
        detail["statusTo"] = case.status
    session.add(CaseEvent(case=case, event_type=event_type, actor=actor, detail=json.dumps(detail, ensure_ascii=False)))
    session.commit()
    return True


PENDING_WINDOW = timedelta(minutes=30)


def latest_pending_draft(session: Session, actor: str) -> DraftOut | None:
    """이 담당자가 최근 30분 안에 준비했지만 아직 초안함 저장이 기록되지 않은 초안.

    작성 창의 [LEONA 초안 마무리] 버튼이 패널의 쪽지를 못 읽을 때 쓰는 보조 경로다.
    """
    since = datetime.now(timezone.utc) - PENDING_WINDOW
    events = session.scalars(
        select(CaseEvent)
        .where(CaseEvent.actor == actor, CaseEvent.event_type.in_(["DRAFT_PREPARED", "DRAFT_SAVED"]))
        .order_by(CaseEvent.id.desc())
        .limit(1)
    ).all()
    if not events or events[0].event_type != "DRAFT_PREPARED":
        return None
    event = events[0]
    created = event.created_at if event.created_at.tzinfo else event.created_at.replace(tzinfo=timezone.utc)
    if created < since:
        return None

    detail = json.loads(event.detail or "{}")
    case = event.case
    attachments = []
    if detail.get("kind") == "QUOTE":
        attachments = [
            DraftAttachment(filename=f.name, url=f"/api/cases/{case.case_id}/quotes/{f.name}", size=f.stat().st_size)
            for f in quote_files(case)
        ]
    return DraftOut(
        kind=detail.get("kind", "SUPPLEMENT"),
        case_id=case.case_id,
        to=_recipient(case),
        subject=detail.get("subject", ""),
        html_body="",
        attachments=attachments,
    )


def quote_files(case: Case) -> list[Path]:
    """C트랙이 만든 견적서(PDF·XLSX)는 storage/<케이스ID>/quotes/ 에 둔다."""
    folder = settings.storage_dir / case.case_id / "quotes"
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.iterdir() if p.suffix.lower() in QUOTE_EXTENSIONS)


def save_quote_file(case: Case, filename: str, data: bytes) -> str:
    if Path(filename).suffix.lower() not in QUOTE_EXTENSIONS:
        raise DraftError("견적서는 PDF 또는 XLSX 파일만 올릴 수 있습니다.")
    return storage.save_file(f"{case.case_id}/quotes", filename, data)


def quote_draft(session: Session, case: Case, actor: str | None) -> DraftOut:
    """견적서 송부 초안. 견적서 PDF·XLSX를 첨부한다 (FR-505)."""
    files = quote_files(case)
    if not files:
        raise DraftError(
            "이 케이스에는 아직 견적서가 없어 송부 초안을 만들지 않았습니다. "
            "견적서는 필수 항목이 모두 확보되고 요율이 등록된 구간이면 자동으로 만들어집니다 "
            "(케이스 상태 '계산완료'). 상태가 '정보부족'이면 먼저 보완 요청을, '보류'면 사유를 확인하세요."
        )

    subject = f"[{case.case_id}] 견적서 송부 / Quotation - {_original(case).subject or ''}".strip()
    names = "".join(f"<li>{escape(f.name)}</li>" for f in files)
    body = (
        "<p>안녕하세요, 레오나해운항공㈜입니다.<br>"
        "요청하신 견적서를 첨부와 같이 송부드립니다. 검토 부탁드립니다.</p>"
        f"<ul>{names}</ul>"
        f"<p>감사합니다.<br>{SIGNATURE_KO}</p>"
        "<hr>"
        "<p>Hello, please find the attached quotation for your request.<br>"
        f"Thank you.<br>{SIGNATURE_EN}</p>"
    )
    attachments = [
        DraftAttachment(filename=f.name, url=f"/api/cases/{case.case_id}/quotes/{f.name}", size=f.stat().st_size)
        for f in files
    ]
    _record(session, case, "QUOTE", actor, {"subject": subject, "files": [f.name for f in files]})
    return DraftOut(kind="QUOTE", case_id=case.case_id, to=_recipient(case), subject=subject, html_body=body, attachments=attachments)
