""".eml(메일 원문) → MailSnapshot 변환. 애드인 원문과 .eml 파일이 같은 파서를 쓴다.

외부 라이브러리 없이 파이썬 표준 email 모듈만 사용한다.
"""

import hashlib
import re
from dataclasses import dataclass, field
from datetime import datetime
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import getaddresses, parsedate_to_datetime
from html.parser import HTMLParser

from .schemas import Address, AttachmentInfo, MailSnapshot


@dataclass
class AttachmentFile:
    filename: str
    content_type: str
    data: bytes


@dataclass
class ParsedMail:
    snapshot: MailSnapshot
    attachments: list[AttachmentFile] = field(default_factory=list)


def parse_eml(raw: bytes) -> ParsedMail:
    msg: EmailMessage = BytesParser(policy=policy.default).parsebytes(raw)  # type: ignore[assignment]

    attachments = _attachments(msg)
    senders = _addresses(msg, "from")
    snapshot = MailSnapshot(
        subject=_header(msg, "subject"),
        from_=senders[0] if senders else Address(),
        to=_addresses(msg, "to"),
        cc=_addresses(msg, "cc"),
        received_at=_received_at(msg),
        body_text=_body_text(msg),
        internet_message_id=_header(msg, "message-id") or None,
        in_reply_to=_header(msg, "in-reply-to") or None,
        references=_header(msg, "references").split(),
        attachments=[
            AttachmentInfo(
                filename=a.filename,
                content_type=a.content_type,
                size=len(a.data),
                sha256=hashlib.sha256(a.data).hexdigest(),
            )
            for a in attachments
        ],
    )
    return ParsedMail(snapshot=snapshot, attachments=attachments)


def _header(msg: EmailMessage, name: str) -> str:
    try:
        value = msg.get(name)
    except Exception:  # 형식이 깨진 헤더는 빈 값으로 처리
        return ""
    return " ".join(str(value).split()) if value is not None else ""


def _addresses(msg: EmailMessage, name: str) -> list[Address]:
    try:
        values = msg.get_all(name) or []
    except Exception:
        return []
    return [Address(name=n, email=e) for n, e in getaddresses([str(v) for v in values]) if e]


def _received_at(msg: EmailMessage) -> datetime | None:
    """수신 시각: 가장 위(마지막으로 거친 서버)의 Received 헤더 시각, 없으면 Date(발송 시각)."""
    candidates = []
    received = msg.get_all("received") or []
    if received:
        candidates.append(str(received[0]).rsplit(";", 1)[-1])
    candidates.append(_header(msg, "date"))
    for text in candidates:
        try:
            return parsedate_to_datetime(text.strip())
        except (TypeError, ValueError, IndexError):
            continue
    return None


def _body_text(msg: EmailMessage) -> str:
    """본문: text/plain이 있으면 그대로, HTML만 있으면 태그를 걷어 글자만 남긴다."""
    part = msg.get_body(preferencelist=("plain", "html"))
    if part is None:
        return ""
    text = _decode_text(part)
    if part.get_content_type() == "text/html":
        text = html_to_text(text)
    return text.strip()


def _decode_text(part: EmailMessage) -> str:
    try:
        return part.get_content()
    except (LookupError, UnicodeDecodeError):
        # 알 수 없는 문자셋(ks_c_5601-1987 등 변형 표기) → 한국어 메일에서 가장 흔한 cp949로 재시도
        payload = part.get_payload(decode=True) or b""
        for charset in (part.get_content_charset(), "cp949", "utf-8"):
            try:
                return payload.decode(charset or "utf-8")
            except (LookupError, UnicodeDecodeError):
                continue
        return payload.decode("utf-8", errors="replace")


def _attachments(msg: EmailMessage) -> list[AttachmentFile]:
    files = []
    for part in msg.iter_attachments():
        filename = part.get_filename()
        if not filename and part.get_content_maintype() != "message":
            continue  # 파일명 없는 본문 삽입 이미지 등은 첨부로 보지 않는다
        if part.get_content_type() == "message/rfc822":
            inner = part.get_payload(0) if part.is_multipart() else part.get_content()
            data = inner.as_bytes()
            filename = filename or "attached-message.eml"
        else:
            data = part.get_payload(decode=True) or b""
        files.append(AttachmentFile(filename=filename, content_type=part.get_content_type(), data=data))
    return files


class _TextExtractor(HTMLParser):
    BLOCK_TAGS = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "table"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self._skip += 1
        elif tag in self.BLOCK_TAGS:
            self.parts.append("\n")
        elif tag in ("td", "th"):
            self.parts.append("\t")  # 표의 칸 구분을 남겨 B트랙이 표 형태를 알 수 있게 한다

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        if not self._skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    parser = _TextExtractor()
    parser.feed(html)
    lines = "".join(parser.parts).splitlines()
    # 줄마다 연속 공백(nbsp 포함)을 한 칸으로 줄이고, 표 칸 구분(\t)은 남긴다
    cleaned = [re.sub(r"[  ]+", " ", line).strip(" ") for line in lines]
    return "\n".join(line for line in cleaned if line.strip())
