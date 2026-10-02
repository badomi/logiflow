"""케이스 생성·회신 병합 뒤에 실행되는 처리 단계 — B트랙(추출)·C트랙(검증) 연결 지점.

A트랙은 "언제 부르고, 얼마나 걸렸고, 결과가 어땠는지"를 책임진다 (FR-101 수용 기준: 추출 수행, 60초 이내).
실제 추출(로컬 LLM)과 검증(룰 베이스)은 B·C트랙이 아래 Extractor / Validator 규격에 맞춰 구현해 register()로 등록한다.
등록 전에는 아무것도 하지 않는 기본 구현이 돌아가며, 결과는 '미연결'로 남는다.

흐름:
  start()   — 요청 안에서 즉시: 상태를 파싱중/재파싱으로 바꾸고 PIPELINE_STARTED 기록 (담당자는 기다리지 않는다)
  execute() — 응답 뒤 백그라운드에서: 추출 → 검증 실행, 걸린 시간·결과를 PIPELINE_RUN(실패 시 PIPELINE_FAILED)으로 기록
"""

import json
import logging
import time
from dataclasses import dataclass
from typing import Protocol

from sqlalchemy.orm import Session

from .models import Case, CaseEvent

log = logging.getLogger(__name__)

# FR-101(변경 승인본)·NFR-02: 케이스 생성 실행 후 완료까지 60초 이내
TIME_LIMIT_MS = 60_000

# 추출이 도는 동안의 상태 (6장 caseStatus 값)
RUNNING_STATUS = {"CASE_CREATED": "파싱중", "REPLY_MERGED": "재파싱", "MANUAL_RERUN": "재파싱", "FIELD_CORRECTED": "재파싱"}


class Extractor(Protocol):
    """B트랙: 케이스의 메일(원문·회신·첨부)을 읽어 6장 필드 JSON을 돌려준다."""

    name: str

    def extract(self, session: Session, case: Case) -> dict | None: ...


class Validator(Protocol):
    """C트랙: 추출 결과로 누락·정합성을 판정해 결과 JSON을 돌려준다. 케이스 상태(정보부족 등)도 여기서 정한다."""

    name: str

    def validate(self, session: Session, case: Case, extraction: dict | None) -> dict | None: ...


class _NotConnected:
    name = "not-connected"

    def extract(self, session, case):
        return None

    def validate(self, session, case, extraction):
        return None


_extractor: Extractor = _NotConnected()
_validator: Validator = _NotConnected()


def register(extractor: Extractor | None = None, validator: Validator | None = None) -> None:
    global _extractor, _validator
    if extractor is not None:
        _extractor = extractor
    if validator is not None:
        _validator = validator


@dataclass
class Job:
    """백그라운드에서 실행할 추출 작업 한 건"""

    case_pk: int
    trigger: str  # CASE_CREATED / REPLY_MERGED / MANUAL_RERUN(다시 추출) / FIELD_CORRECTED(수동 보정)
    actor: str | None
    previous_status: str


def start(session: Session, case: Case, trigger: str, actor: str | None) -> Job:
    """추출 시작을 기록하고 상태를 '파싱중/재파싱'으로 바꾼다. 실제 실행은 execute()."""
    previous = case.status
    case.status = RUNNING_STATUS.get(trigger, "파싱중")
    _event(session, case, "PIPELINE_STARTED", actor, {"trigger": trigger, "statusFrom": previous})
    session.commit()
    return Job(case_pk=case.id, trigger=trigger, actor=actor, previous_status=previous)


def execute(job: Job) -> None:
    """추출 → 검증을 실행하고 결과·걸린 시간을 기록한다. 새 DB 세션을 연다(응답이 끝난 뒤 실행되므로).

    실패해도 케이스는 지워지지 않고 상태 '실패'와 이력이 남아 다시 실행할 수 있다 (NFR-03).
    """
    from .db import SessionLocal

    with SessionLocal() as session:
        case = session.get(Case, job.case_pk)
        if case is None:
            return
        running = case.status
        detail: dict = {"trigger": job.trigger, "extractor": _extractor.name, "validator": _validator.name}
        began = time.monotonic()
        try:
            extraction = _extractor.extract(session, case)
            validation = _validator.validate(session, case, extraction)
            event_type = "PIPELINE_RUN"
            detail["extraction"] = extraction
            detail["validation"] = validation
            # 추출기가 아직 없거나 검증이 상태를 정하지 않았으면 원래 상태로 되돌린다 (파싱중에 멈춰 있지 않게)
            if case.status == running:
                case.status = job.previous_status
        except Exception as error:  # B·C 코드의 오류가 A트랙 흐름(케이스 저장)을 깨지 않게 한다
            log.exception("pipeline failed for %s", case.case_id)
            session.rollback()
            case = session.get(Case, job.case_pk)
            event_type = "PIPELINE_FAILED"
            detail["error"] = str(error)
            case.status = "실패"

        elapsed_ms = int((time.monotonic() - began) * 1000)
        detail["elapsedMs"] = elapsed_ms
        detail["withinLimit"] = elapsed_ms <= TIME_LIMIT_MS
        detail["statusTo"] = case.status
        _event(session, case, event_type, job.actor, detail)
        session.commit()


def run(session: Session, case: Case, trigger: str, actor: str | None) -> None:
    """start + execute를 바로 이어서 실행 (.eml 리더 일괄 등록처럼 기다려도 되는 경우)."""
    execute(start(session, case, trigger, actor))
    session.expire_all()


def _event(session: Session, case: Case, event_type: str, actor: str | None, detail: dict) -> None:
    session.add(
        CaseEvent(case=case, event_type=event_type, actor=actor, detail=json.dumps(detail, ensure_ascii=False, default=str))
    )
