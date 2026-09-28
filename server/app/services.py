"""A트랙 업무 로직: 케이스 생성(FR-101·102·103·105), 회신 식별·병합(FR-104·207)."""

import base64
import binascii
import json
from datetime import datetime, timezone

from sqlalchemy import select
from sqlalchemy.orm import Session

from . import pipeline, storage
from .case_id import find_case_ids, issue_case_id
from .mail_parser import AttachmentFile, ParsedMail, parse_eml
from .models import Attachment, Case, CaseEvent, Mail
from .schemas import (
    Address,
    AttachmentInfo,
    CaseDetail,
    CaseSummary,
    EventOut,
    MailInput,
    MailOut,
    MailSnapshot,
    ReplyCandidate,
    ReplyMatchResult,
)


class InputError(ValueError):
    """요청 내용이 잘못됨 (HTTP 400)"""


class NotFoundError(LookupError):
    """케이스 없음 (HTTP 404)"""


class ConflictError(RuntimeError):
    """이미 다른 케이스에 등록된 메일 (HTTP 409)"""


# ---------------------------------------------------------------- 입력 해석


def resolve_input(data: MailInput) -> tuple[ParsedMail, bytes | None, str]:
    """애드인 요청 → (파싱 결과, 원문 바이트, 출처).

    원문(.eml)이 있으면 서버 파서 결과를 기준으로 하고, .eml에 없는 값(conversationId)과
    Outlook이 더 정확히 아는 값(수신 시각)은 애드인이 보낸 snapshot으로 채운다.
    """
    if data.eml_base64:
        try:
            raw = base64.b64decode(data.eml_base64, validate=True)
        except (binascii.Error, ValueError) as error:
            raise InputError("emlBase64가 올바른 Base64가 아닙니다.") from error
        parsed = parse_eml(raw)
        if data.snapshot:
            hint = data.snapshot
            parsed.snapshot.conversation_id = hint.conversation_id or parsed.snapshot.conversation_id
            parsed.snapshot.received_at = hint.received_at or parsed.snapshot.received_at
            parsed.snapshot.internet_message_id = parsed.snapshot.internet_message_id or hint.internet_message_id
        return parsed, raw, "ADDIN"
    if data.snapshot:
        return ParsedMail(snapshot=data.snapshot), None, "ADDIN"
    raise InputError("emlBase64 또는 snapshot 중 하나는 있어야 합니다.")


def normalize_message_id(value: str | None) -> str | None:
    """Message-ID 비교용 정규화: 공백 제거, 꺾쇠(<>) 통일."""
    if not value or not value.strip():
        return None
    value = value.strip()
    return value if value.startswith("<") else f"<{value.strip('<>')}>"


# ---------------------------------------------------------------- 케이스 생성


def find_registered_mail(session: Session, internet_message_id: str | None) -> Mail | None:
    message_id = normalize_message_id(internet_message_id)
    if message_id is None:
        return None
    return session.scalar(select(Mail).where(Mail.internet_message_id == message_id))


def create_case(
    session: Session, parsed: ParsedMail, raw: bytes | None, source: str, actor: str | None
) -> tuple[Case, bool]:
    """메일 1통으로 케이스를 만든다. 이미 등록된 메일이면 기존 케이스를 돌려준다 (duplicate=True)."""
    existing = find_registered_mail(session, parsed.snapshot.internet_message_id)
    if existing is not None:
        return existing.case, True

    case = Case(case_id=issue_case_id(session), status="수신", created_by=actor)
    session.add(case)
    _add_mail(session, case, "ORIGINAL", parsed, raw, source)
    _add_event(session, case, "CASE_CREATED", actor, {"source": source, "subject": parsed.snapshot.subject})
    session.commit()

    pipeline.run(session, case, trigger="CASE_CREATED", actor=actor)
    return case, False


def _add_mail(session: Session, case: Case, role: str, parsed: ParsedMail, raw: bytes | None, source: str) -> Mail:
    snap = parsed.snapshot
    index = len(case.mails) + 1
    folder = f"{case.case_id}/mail-{index:02d}"

    mail = Mail(
        case=case,
        role=role,
        source=source,
        internet_message_id=normalize_message_id(snap.internet_message_id),
        conversation_id=snap.conversation_id,
        in_reply_to=normalize_message_id(snap.in_reply_to),
        ref_headers=" ".join(filter(None, map(normalize_message_id, snap.references))) or None,
        subject=snap.subject[:1000],
        from_name=snap.from_.name,
        from_email=snap.from_.email,
        to_addrs=_dump_addresses(snap.to),
        cc_addrs=_dump_addresses(snap.cc),
        received_at=snap.received_at.astimezone(timezone.utc) if snap.received_at else None,
        body_text=snap.body_text,
        raw_path=storage.save_file(folder, "original.eml", raw) if raw else None,
    )
    session.add(mail)
    for file in parsed.attachments:
        _add_attachment(mail, folder, file, snap.attachments)
    return mail


def _add_attachment(mail: Mail, folder: str, file: AttachmentFile, infos: list[AttachmentInfo]) -> None:
    info = next((i for i in infos if i.filename == file.filename and i.size == len(file.data)), None)
    mail.attachments.append(
        Attachment(
            filename=file.filename[:500],
            content_type=file.content_type,
            size_bytes=len(file.data),
            sha256=info.sha256 if info else "",
            storage_path=storage.save_file(f"{folder}/attachments", file.filename, file.data),
        )
    )


def _add_event(session: Session, case: Case, event_type: str, actor: str | None, detail: dict) -> None:
    session.add(CaseEvent(case=case, event_type=event_type, actor=actor, detail=json.dumps(detail, ensure_ascii=False)))


def _dump_addresses(addresses: list[Address]) -> str:
    return json.dumps([a.model_dump() for a in addresses], ensure_ascii=False)


# ---------------------------------------------------------------- 회신 식별·병합


def match_reply(session: Session, snapshot: MailSnapshot) -> ReplyMatchResult:
    """회신 메일이 어느 케이스에 속하는지 후보를 찾는다 (FR-104). 병합은 담당자가 확인한 뒤에 한다.

    근거가 강한 순서: 제목의 케이스 ID(1.0) > 메일 헤더 In-Reply-To·References(0.9) > Outlook 대화 ID(0.7)
    """
    registered = find_registered_mail(session, snapshot.internet_message_id)
    candidates: dict[str, ReplyCandidate] = {}

    def add(case: Case, score: float, reason: str) -> None:
        found = candidates.get(case.case_id)
        if found is None:
            candidates[case.case_id] = ReplyCandidate(
                case_id=case.case_id, subject=_case_subject(case), score=score, reasons=[reason]
            )
        elif reason not in found.reasons:
            found.reasons.append(reason)
            found.score = max(found.score, score)

    for case_id in find_case_ids(snapshot.subject):
        case = session.scalar(select(Case).where(Case.case_id == case_id))
        if case is not None:
            add(case, 1.0, "SUBJECT_CASE_ID")

    header_ids = {normalize_message_id(v) for v in [snapshot.in_reply_to, *snapshot.references]} - {None}
    if header_ids:
        for mail in session.scalars(select(Mail).where(Mail.internet_message_id.in_(header_ids))):
            add(mail.case, 0.9, "HEADER")

    if snapshot.conversation_id:
        for mail in session.scalars(select(Mail).where(Mail.conversation_id == snapshot.conversation_id)):
            add(mail.case, 0.7, "CONVERSATION")

    if registered is not None:
        candidates.pop(registered.case.case_id, None)  # 자기 자신이 속한 케이스는 후보에서 뺀다

    ranked = sorted(candidates.values(), key=lambda c: (-c.score, -len(c.reasons), c.case_id))
    return ReplyMatchResult(
        already_registered_case_id=registered.case.case_id if registered else None,
        candidates=ranked,
    )


def merge_reply(
    session: Session, case_id: str, parsed: ParsedMail, raw: bytes | None, source: str, actor: str | None
) -> Case:
    """담당자가 확인한 케이스에 회신을 합치고, 추출→검증을 다시 실행한다 (FR-104·FR-207)."""
    case = get_case(session, case_id)
    existing = find_registered_mail(session, parsed.snapshot.internet_message_id)
    if existing is not None:
        if existing.case_pk == case.id:
            return case  # 이미 이 케이스에 합쳐진 회신 — 다시 눌러도 결과가 같다
        raise ConflictError(f"이 메일은 이미 {existing.case.case_id} 케이스에 등록되어 있습니다.")

    _add_mail(session, case, "REPLY", parsed, raw, source)
    _add_event(session, case, "REPLY_MERGED", actor, {"subject": parsed.snapshot.subject})
    session.commit()

    pipeline.run(session, case, trigger="REPLY_MERGED", actor=actor)
    return case


# ---------------------------------------------------------------- 조회


def get_case(session: Session, case_id: str) -> Case:
    case = session.scalar(select(Case).where(Case.case_id == case_id))
    if case is None:
        raise NotFoundError(f"케이스 {case_id}를 찾을 수 없습니다.")
    return case


def list_cases(session: Session, limit: int) -> list[CaseSummary]:
    cases = session.scalars(select(Case).order_by(Case.id.desc()).limit(limit))
    return [_summary(c) for c in cases]


def case_detail(case: Case) -> CaseDetail:
    return CaseDetail(
        **_summary(case).model_dump(),
        mails=[
            MailOut(role=m.role, source=m.source, snapshot=mail_to_snapshot(m), has_raw=m.raw_path is not None)
            for m in case.mails
        ],
        events=[
            EventOut(
                event_type=e.event_type,
                actor=e.actor,
                detail=json.loads(e.detail) if e.detail else None,
                created_at=_as_utc(e.created_at),
            )
            for e in case.events
        ],
    )


def mail_to_snapshot(mail: Mail) -> MailSnapshot:
    return MailSnapshot(
        subject=mail.subject or "",
        from_=Address(name=mail.from_name or "", email=mail.from_email or ""),
        to=[Address(**a) for a in json.loads(mail.to_addrs or "[]")],
        cc=[Address(**a) for a in json.loads(mail.cc_addrs or "[]")],
        received_at=_as_utc(mail.received_at),
        body_text=mail.body_text or "",
        conversation_id=mail.conversation_id,
        internet_message_id=mail.internet_message_id,
        in_reply_to=mail.in_reply_to,
        references=(mail.ref_headers or "").split(),
        attachments=[
            AttachmentInfo(filename=a.filename, content_type=a.content_type, size=a.size_bytes, sha256=a.sha256)
            for a in mail.attachments
        ],
    )


def _summary(case: Case) -> CaseSummary:
    first = case.mails[0] if case.mails else None
    return CaseSummary(
        case_id=case.case_id,
        status=case.status,
        subject=(first.subject or "") if first else "",
        from_email=(first.from_email or "") if first else "",
        created_at=_as_utc(case.created_at),
    )


def _case_subject(case: Case) -> str:
    return (case.mails[0].subject or "") if case.mails else ""


def _as_utc(value: datetime | None) -> datetime | None:
    """SQLite는 시간대 정보를 버리고 저장하므로, 읽을 때 UTC로 다시 붙인다."""
    if value is None or value.tzinfo is not None:
        return value
    return value.replace(tzinfo=timezone.utc)
