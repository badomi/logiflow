"""메일 초안 내용 만들기 — 보완 요청(FR-304·305·510), 견적서 송부(FR-505).

A트랙은 "받는 사람·제목·본문·첨부"를 만들어 패널에 넘기고, 패널이 Outlook 작성 창을 띄운다.
발송은 항상 담당자가 직접 누른다. 이 모듈에는 메일을 보내는 코드가 없다 (정의서 12장 설계 고정).
"""

import json
from html import escape
from pathlib import Path

from sqlalchemy.orm import Session

from . import storage
from .config import settings
from .models import Case, CaseEvent
from .schemas import Address, DraftAttachment, DraftOut

SIGNATURE_KO = "LEONA Shipping & Air 견적 담당"
SIGNATURE_EN = "LEONA Shipping & Air Quotation Team"

QUOTE_EXTENSIONS = {".pdf", ".xlsx"}


class DraftError(ValueError):
    """초안을 만들 수 없음 (문항·견적서 없음 등)"""


def _original(case: Case):
    if not case.mails:
        raise DraftError("케이스에 메일이 없습니다.")
    return case.mails[0]


def _recipient(case: Case) -> list[Address]:
    """받는 사람: 최초 요청 메일의 발신자(화주). FR-208(회신 주소 추출)이 연결되면 그 값을 우선한다."""
    original = _original(case)
    return [Address(name=original.from_name or "", email=original.from_email or "")]


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
    if not questions:
        raise DraftError("보완 요청 문항이 없습니다.")

    subject = f"[{case.case_id}] 견적 보완 요청 / Request for additional information - {_original(case).subject or ''}".strip()
    items = "".join(f"<li>{escape(q)}</li>" for q in questions)
    body = (
        "<p>안녕하세요, LEONA Shipping &amp; Air입니다.<br>"
        "요청하신 견적을 산출하기 위해 아래 항목을 추가로 알려 주시기 바랍니다.</p>"
        f"<ol>{items}</ol>"
        f"<p>회신 시 제목의 케이스 번호({escape(case.case_id)})를 유지해 주시면 빠르게 처리됩니다.<br>"
        f"감사합니다.<br>{SIGNATURE_KO}</p>"
        "<hr>"
        "<p>Hello, this is LEONA Shipping &amp; Air.<br>"
        "To prepare your quotation, please provide the information listed above.<br>"
        f"Please keep the case number ({escape(case.case_id)}) in the subject when replying.<br>"
        f"Thank you.<br>{SIGNATURE_EN}</p>"
    )
    _record(session, case, "SUPPLEMENT", actor, {"subject": subject, "questionCount": len(questions)})
    return DraftOut(kind="SUPPLEMENT", to=_recipient(case), subject=subject, html_body=body, attachments=[])


MAIL_EVENT_TYPES = {"DRAFT_SAVED", "MAIL_SENT"}

# 담당자가 메일을 보낸 뒤의 케이스 상태 (6장 caseStatus). 상태 전이 규칙의 최종 결정은 D트랙.
STATUS_AFTER_SENT = {"SUPPLEMENT": "보완대기", "QUOTE": "발송완료"}


def record_mail_event(session: Session, case: Case, event_type: str, kind: str | None, subject: str, actor: str | None):
    """작성 창(초안 저장)·보내기 이벤트를 이력에 남긴다. 기록 시각이 곧 발송 시각이다 (FR-505).

    이 함수는 기록만 한다. 메일을 보내는 것은 Outlook에서 담당자가 직접 누른 보내기 버튼이다.
    """
    if event_type not in MAIL_EVENT_TYPES:
        raise DraftError(f"알 수 없는 이벤트입니다: {event_type}")
    if kind is None:
        kind = "QUOTE" if "견적서 송부" in subject else "SUPPLEMENT" if "보완 요청" in subject else None

    detail = {"kind": kind, "subject": subject}
    if event_type == "MAIL_SENT" and kind in STATUS_AFTER_SENT:
        detail["statusFrom"] = case.status
        case.status = STATUS_AFTER_SENT[kind]
        detail["statusTo"] = case.status
    session.add(CaseEvent(case=case, event_type=event_type, actor=actor, detail=json.dumps(detail, ensure_ascii=False)))
    session.commit()


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
        raise DraftError("첨부할 견적서가 없습니다. 견적서(PDF·XLSX) 생성 후 다시 시도하세요.")

    subject = f"[{case.case_id}] 견적서 송부 / Quotation - {_original(case).subject or ''}".strip()
    names = "".join(f"<li>{escape(f.name)}</li>" for f in files)
    body = (
        "<p>안녕하세요, LEONA Shipping &amp; Air입니다.<br>"
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
    return DraftOut(kind="QUOTE", to=_recipient(case), subject=subject, html_body=body, attachments=attachments)
