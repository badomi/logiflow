"""pipeline.Extractor 구현 — 케이스의 메일·첨부를 엔진에 넣고, 결과를 quote_inputs·extraction_fields에 저장한다.

  - 임계치 이상(filled)인 값만 quote_inputs에 쓴다. review·missing은 쓰지 않아 검증 단계에서 보완 요청 대상이 된다.
  - 담당자가 수정한 필드(method='manual')는 재추출이 덮어쓰지 않는다 (FR-206).
  - 돌려주는 dict는 A트랙이 case_events(PIPELINE_RUN)에 그대로 남긴다 (NFR-05).
"""

import json

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from .. import storage
from ..config import settings
from ..llm import OllamaClient
from ..models import INPUT_FIELDS, Case, CaseEvent, ExtractionField, QuoteInput
from . import attachments, engine
from .base import MailText

INT_FIELDS = {"qty", "boxL", "boxW", "boxH"}
FLOAT_FIELDS = {"totalCbm", "grossWeightKg", "invoiceValue"}


def mails_of(case: Case) -> list[MailText]:
    out = []
    for mail in case.mails:
        files = []
        for att in mail.attachments:
            try:
                files.append((att.filename, attachments.extract_text(att.filename, storage.read_file(att.storage_path))))
            except (OSError, ValueError):
                continue
        out.append(MailText(mail.subject or "", mail.from_name or "", mail.from_email or "", mail.body_text or "",
                            mail.role, files))
    return out


def asked_fields(case: Case) -> list[str | None] | None:
    """마지막 보완 요청에서 보낸 질문의 항목 (번호 순). 번호로만 답한 회신을 읽는 데 쓴다."""
    for event in reversed(case.events):
        if event.event_type == "DRAFT_PREPARED":
            detail = json.loads(event.detail or "{}")
            if detail.get("kind") == "SUPPLEMENT":
                return detail.get("questionFields")
    return None


def manual_fields(session: Session, case: Case) -> set[str]:
    rows = session.scalars(
        select(ExtractionField.field_name).where(ExtractionField.case_pk == case.id, ExtractionField.method == "manual")
    ).all()
    return set(rows)


class HybridExtractor:
    """룰 + 로컬 LLM 하이브리드. llm_mode: 'hybrid'(LLM 사용) / 'rules'(룰만)."""

    def __init__(self, llm=None, llm_mode: str = "hybrid"):
        self._llm = llm
        self.llm_mode = llm_mode
        self.name = f"hybrid-{engine.SCHEMA_VERSION}" if llm_mode == "hybrid" else f"rules-{engine.SCHEMA_VERSION}"

    def _client(self):
        if self.llm_mode != "hybrid":
            return None
        if self._llm is None:
            self._llm = OllamaClient()
        return self._llm

    def extract(self, session: Session, case: Case) -> dict:
        try:
            client = self._client()
        except Exception as error:  # 주소가 외부(NFR-04)이거나 잘못됨 → 룰만으로 진행하고 사유를 남긴다
            client, setup_error = None, str(error)
        else:
            setup_error = None
        result = engine.run(mails_of(case), llm=client, model_name=settings.llm_model if client else None,
                            asked=asked_fields(case), scope=settings.llm_scope)
        if setup_error:
            result["llm"]["error"] = setup_error

        locked = manual_fields(session, case)
        record = case.quote_input or QuoteInput(case=case)
        changed = []
        for name, value in engine.saved_values(result).items():
            if name in locked:
                continue
            if getattr(record, name) != value:
                setattr(record, name, value)
                changed.append(name)
        session.add(record)

        run_no = (session.scalar(
            select(func.max(ExtractionField.run_no)).where(ExtractionField.case_pk == case.id)) or 0) + 1
        for name in INPUT_FIELDS:
            f = result["fields"][name]
            session.add(ExtractionField(
                case=case, run_no=run_no, field_name=name,
                value=None if f["value"] is None else json.dumps(f["value"], ensure_ascii=False),
                score=f["score"], evidence=f["evidence"], origin=f["origin"], status=f["status"],
                # 해석 못 한 값은 method에 표시 → 'python -m app.extraction.report'로 모아 보고 사전을 보강한다
                method=(f"{f['method']}:unparsed" if f.get("reason") == "UNPARSED" and f["method"] else f["method"]),
            ))
        session.flush()

        result["runNo"] = run_no
        result["changed"] = changed
        result["lockedByManual"] = sorted(locked)
        return result


# ------------------------------------------------------------------ FR-206 수동 보정


class CorrectionError(ValueError):
    pass


def coerce(field: str, value):
    if field not in INPUT_FIELDS:
        raise CorrectionError(f"6장 필드가 아닙니다: {field}")
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        if field in INT_FIELDS:
            return int(str(value).replace(",", ""))
        if field in FLOAT_FIELDS:
            return float(str(value).replace(",", ""))
    except ValueError as error:
        raise CorrectionError(f"{field}는 숫자여야 합니다: {value}") from error
    return str(value).strip()


def correct_field(session: Session, case: Case, field: str, value, actor: str | None) -> dict:
    """담당자 수정: quote_inputs 갱신 + 수정 이력(extraction_fields method=manual) + 케이스 이력."""
    new = coerce(field, value)
    record = case.quote_input or QuoteInput(case=case)
    before = getattr(record, field)
    setattr(record, field, new)
    session.add(record)
    run_no = session.scalar(select(func.max(ExtractionField.run_no)).where(ExtractionField.case_pk == case.id)) or 0
    session.add(ExtractionField(
        case=case, run_no=run_no, field_name=field,
        value=None if new is None else json.dumps(new, ensure_ascii=False), score=1.0,
        evidence="담당자 수정", origin="담당자", method="manual", status="filled" if new is not None else "missing",
        actor=actor,
    ))
    detail = {"field": field, "before": before, "after": new}
    session.add(CaseEvent(case=case, event_type="FIELD_CORRECTED", actor=actor,
                          detail=json.dumps(detail, ensure_ascii=False, default=str)))
    session.flush()
    return detail


def field_view(session: Session, case: Case) -> list[dict]:
    """케이스 화면용: 필드별 현재 값 + 가장 최근 근거·점수·상태."""
    latest: dict[str, ExtractionField] = {}
    for row in case.extraction_fields:  # id 순이라 마지막이 최신. 담당자 수정(manual)은 이후 추출보다 우선
        if latest.get(row.field_name) is not None and latest[row.field_name].method == "manual" and row.method != "manual":
            continue
        latest[row.field_name] = row
    record = case.quote_input
    out = []
    for name in INPUT_FIELDS:
        row = latest.get(name)
        out.append({
            "field": name,
            "value": getattr(record, name) if record else None,
            "status": row.status if row else "missing",
            "score": row.score if row else 0.0,
            "evidence": row.evidence if row else None,
            "origin": row.origin if row else None,
            "method": row.method if row else None,
            "candidate": json.loads(row.value) if row and row.value and row.status == "review" else None,
        })
    return out
