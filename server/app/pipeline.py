"""케이스 생성·회신 병합 뒤에 실행되는 처리 단계 — B트랙(추출)·C트랙(검증) 연결 지점.

A트랙은 "언제 부르는지"만 책임진다. 실제 추출(로컬 LLM)과 검증(룰 베이스)은
B·C트랙이 아래 Extractor / Validator 규격에 맞춰 구현해 register()로 등록한다.
등록 전에는 아무것도 하지 않는 기본 구현이 돌아가며, 이력에 '미연결'로 남는다.
"""

import json
import logging
from typing import Protocol

from sqlalchemy.orm import Session

from .models import Case, CaseEvent

log = logging.getLogger(__name__)


class Extractor(Protocol):
    """B트랙: 케이스의 메일(원문·회신·첨부)을 읽어 6장 필드 JSON을 돌려준다."""

    name: str

    def extract(self, session: Session, case: Case) -> dict | None: ...


class Validator(Protocol):
    """C트랙: 추출 결과로 누락·정합성을 판정해 결과 JSON을 돌려준다."""

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


def run(session: Session, case: Case, trigger: str, actor: str | None) -> None:
    """추출 → 검증을 차례로 실행한다.

    trigger: CASE_CREATED(케이스 생성, FR-101) / REPLY_MERGED(회신 병합 후 재추출·재검증, FR-207)
    실패해도 케이스는 지워지지 않고 이력에 실패가 남아 다시 실행할 수 있다 (NFR-03).
    """
    detail: dict = {"trigger": trigger, "extractor": _extractor.name, "validator": _validator.name}
    event_type = "PIPELINE_RUN"
    try:
        extraction = _extractor.extract(session, case)
        validation = _validator.validate(session, case, extraction)
        detail["extracted"] = extraction is not None
        detail["validated"] = validation is not None
    except Exception as error:  # B·C 코드의 오류가 A트랙 흐름(케이스 저장)을 깨지 않게 한다
        log.exception("pipeline failed for %s", case.case_id)
        event_type = "PIPELINE_FAILED"
        detail["error"] = str(error)

    session.add(CaseEvent(case=case, event_type=event_type, actor=actor, detail=json.dumps(detail, ensure_ascii=False)))
    session.commit()
