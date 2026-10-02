"""메일 본문 전처리 — 인용문 분리(역할분담서·WBS 3.5 B트랙), 서명 분리(FR-208), 줄 정리.

A트랙은 원문을 자르지 않고 넘긴다(docs/트랙간_전달사항.md 2-1). 여기서:
  - 회신 아래 붙은 이전 메일(인용문)을 떼어낸다 → 이전 값이 '새 값'처럼 다시 뽑히지 않게
  - Outlook HTML 목록 변환으로 '1.' 과 내용이 갈라진 줄을 합친다
  - 맺음말 아래 서명 블록을 따로 떼어 고객사·담당자 추출에 쓴다
"""

import re
from dataclasses import dataclass

from .parsers import clean

_QUOTE_STARTS = (
    re.compile(r"^_{8,}\s*$"),  # Outlook 웹
    re.compile(r"^-{2,}\s*(Original Message|원본 메시지|Forwarded message|전달된 메시지)\s*-{2,}", re.I),
    re.compile(r"^(보낸 사람|From)\s*:\s*\S", re.I),
    re.compile(r".*님이 작성:\s*$"),  # Gmail 한국어
    re.compile(r"^On .{5,120} wrote:\s*$"),  # Gmail·Apple 영어
    re.compile(r"^\d{4}[.-]\s?\d{1,2}[.-]\s?\d{1,2}.{0,40}(작성|wrote)", re.I),
)
_LONE_MARKER = re.compile(r"^\s*(\d{1,2}[.)]|[-•·*▪◦])\s*$")
_CLOSING = re.compile(
    r"^\s*(감사합니다|고맙습니다|수고하세요|잘 부탁드립니다|.{0,10}드림|.{0,10}올림|"
    r"thanks?( you)?|best( regards)?|kind regards|regards|sincerely|cheers|br)[.!,\s]*$",
    re.I,
)


@dataclass
class MailParts:
    body: str  # 새로 쓴 본문 (인용·서명 제외)
    signature: str  # 맺음말 아래 서명 블록
    quoted: str  # 떼어낸 인용문 (참고용, 추출에 쓰지 않음)


def join_broken_list(lines: list[str]) -> list[str]:
    out: list[str] = []
    i = 0
    while i < len(lines):
        if _LONE_MARKER.match(clean(lines[i])) and i + 1 < len(lines) and lines[i + 1].strip():
            out.append(f"{lines[i].strip()} {lines[i + 1].strip()}")
            i += 2
        else:
            out.append(lines[i])
            i += 1
    return out


def split_mail(text: str) -> MailParts:
    # 원문 표기(㈜ 등)는 보존한다. 전각·특수기호 정규화는 비교·파싱할 때만 (parsers.clean)
    lines = (text or "").replace("\u00a0", " ").replace("\r\n", "\n").replace("\r", "\n").split("\n")

    cut = len(lines)
    for i, line in enumerate(lines):
        if any(p.match(clean(line).strip()) for p in _QUOTE_STARTS):
            cut = i
            break
    new, quoted = lines[:cut], lines[cut:]
    quoted += [line for line in new if line.lstrip().startswith(">")]
    new = [line for line in new if not line.lstrip().startswith(">")]
    new = join_broken_list(new)

    # 맺음말(감사합니다 / Best regards)이 끝부분에 있으면 그 아래를 서명으로 본다 (최대 12줄)
    signature: list[str] = []
    for i in range(len(new) - 1, max(-1, len(new) - 16), -1):
        if _CLOSING.match(clean(new[i]).strip()):
            tail = [line for line in new[i + 1:] if line.strip()]
            if len(tail) <= 12:
                signature = tail
                new = new[: i + 1]
            break

    return MailParts(
        body="\n".join(line.rstrip() for line in new).strip(),
        signature="\n".join(line.strip() for line in signature),
        quoted="\n".join(quoted).strip(),
    )
