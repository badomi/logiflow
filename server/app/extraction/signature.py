"""고객사·담당자·회신 이메일 추출 (FR-208): 메일 서명 블록 + 헤더(From).

- contactEmail: 헤더 From이 가장 확실 (0.95). 서명에만 있는 다른 주소는 0.85.
- customerName: 서명에서 회사 표지(㈜, 주식회사, Co., Ltd., Inc. …)가 있는 줄.
- contactName: 서명에서 직함(과장, 대리, Manager …)과 함께 쓴 이름, 또는 헤더 표시 이름(일반 명칭이면 제외).
"""

import re

from .base import Candidate
from .parsers import clean

_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
_COMPANY = re.compile(
    r"㈜|\(주\)|주식회사|유한회사|\bCo\.?\s*,?\s*Ltd\.?|\bLtd\.?\b|\bInc\.?\b|\bCorp(oration)?\.?\b|\bLLC\b|\bGmbH\b|"
    r"\bLimited\b|\bTrading\b|무역|상사|물산|통상|산업",
    re.I,
)
_KO_TITLES = "사원|주임|대리|과장|차장|부장|팀장|실장|이사|상무|전무|대표|매니저|책임|선임|수석|프로"
_KO_NAME = re.compile(rf"^([가-힣]{{2,4}})\s*({_KO_TITLES})(님)?$|^({_KO_TITLES})\s*([가-힣]{{2,4}})$")
_EN_TITLE = re.compile(r"manager|director|executive|officer|specialist|coordinator|assistant|lead|head|ceo|president|sales|purchasing|buyer", re.I)
_EN_NAME = re.compile(r"^((?:Mr|Ms|Mrs|Miss|Dr)\.?\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?|[A-Z][a-z]+(?:\s+[A-Z]\.)?\s+[A-Z][a-z]+)$")
_GENERIC = re.compile(r"담당자|담당|팀|부서|고객센터|영업|구매|info|sales|admin|support|team|office|noreply|no-reply|shipper", re.I)
_PREFIX = re.compile(r"^(회사|회사명|company|업체명?|상호|이름|성명|name|담당자?)\s*[:：]\s*", re.I)


def _strip_prefix(line: str) -> str:
    return _PREFIX.sub("", line).strip(" -|/")


def extract(signature: str, from_name: str, from_email: str, origin: str, order: int) -> list[Candidate]:
    out: list[Candidate] = []
    lines = [_strip_prefix(ln.strip()) for ln in signature.split("\n") if ln.strip()]

    if from_email:
        out.append(Candidate("contactEmail", from_email.strip(), 0.95, f"From: {from_email}", f"{origin} · 헤더", order, "header"))
    for line in lines:
        for email in _EMAIL.findall(line):
            if email.lower() != (from_email or "").lower():
                out.append(Candidate("contactEmail", email, 0.85, line, f"{origin} · 서명", order, "signature"))

    for i, line in enumerate(lines):
        if _EMAIL.search(line) or re.search(r"\d{2,4}[-.)\s]\d{3,4}[-.\s]\d{4}", line):
            continue  # 연락처 줄
        if _COMPANY.search(line) and len(line) <= 60:
            name = re.split(r"\s*[|/·]\s*", line)[0].strip()
            out.append(Candidate("customerName", name, 0.9, line, f"{origin} · 서명", order, "signature"))
            continue
        ko = _KO_NAME.match(clean(line).replace(" / ", " "))
        if ko:
            name = ko[1] or ko[5]
            out.append(Candidate("contactName", f"{name} {ko[2] or ko[4]}", 0.9, line, f"{origin} · 서명", order, "signature"))
            continue
        en = _EN_NAME.match(line)
        if en:
            near = " ".join(lines[max(0, i - 1): i + 2])
            score = 0.9 if (_EN_TITLE.search(near) or line.startswith(("Mr", "Ms", "Mrs", "Miss", "Dr"))) else 0.7
            out.append(Candidate("contactName", en[1], score, line, f"{origin} · 서명", order, "signature"))

    display = (from_name or "").strip()
    if display and not _GENERIC.search(display) and not _EMAIL.search(display):
        if _COMPANY.search(display):
            out.append(Candidate("customerName", display, 0.8, f"From: {display}", f"{origin} · 헤더", order, "header"))
        else:
            out.append(Candidate("contactName", display, 0.8, f"From: {display}", f"{origin} · 헤더", order, "header"))
    return out
