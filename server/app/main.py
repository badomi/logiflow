"""A트랙 백엔드 API 서버.

실행: server 폴더에서  .venv\\Scripts\\python -m uvicorn app.main:app --reload --port 8000
API 문서(자동 생성): http://localhost:8000/docs
"""

from contextlib import asynccontextmanager

from fastapi import Depends, FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session

from . import services
from .db import get_session, init_db
from .schemas import CaseDetail, CaseSummary, CreateCaseResult, MailInput, ReplyMatchResult


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(title="LEONA 자동견적 — A트랙 API", version="0.1.0", lifespan=lifespan)

# 애드인 개발 서버(https://localhost:3000)에서 직접 부를 때를 위한 허용. 외부 도메인은 허용하지 않는다.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000"],
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


def _http_error(error: Exception) -> HTTPException:
    if isinstance(error, services.NotFoundError):
        return HTTPException(status_code=404, detail=str(error))
    if isinstance(error, services.ConflictError):
        return HTTPException(status_code=409, detail=str(error))
    return HTTPException(status_code=400, detail=str(error))


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/cases", response_model=CreateCaseResult)
def create_case(body: MailInput, session: Session = Depends(get_session)):
    """[케이스 생성] 담당자가 패널에서 누른다 (FR-101). ID 발급(FR-102)·원문 보관(FR-103)·중복 방지(FR-105)."""
    try:
        parsed, raw, source = services.resolve_input(body)
        case, duplicate = services.create_case(session, parsed, raw, source, body.actor)
    except (services.InputError, services.NotFoundError, services.ConflictError) as error:
        raise _http_error(error) from error
    return CreateCaseResult(duplicate=duplicate, case=services.case_detail(case))


@app.get("/api/cases", response_model=list[CaseSummary])
def list_cases(limit: int = Query(50, ge=1, le=200), session: Session = Depends(get_session)):
    """최근 케이스 목록. 회신 자동 식별이 실패했을 때 담당자가 직접 고르는 용도 (FR-104)."""
    return services.list_cases(session, limit)


@app.get("/api/cases/{case_id}", response_model=CaseDetail)
def get_case(case_id: str, session: Session = Depends(get_session)):
    try:
        return services.case_detail(services.get_case(session, case_id))
    except services.NotFoundError as error:
        raise _http_error(error) from error


@app.post("/api/replies/match", response_model=ReplyMatchResult)
def match_reply(body: MailInput, session: Session = Depends(get_session)):
    """[회신 식별] 선택한 메일이 어느 케이스의 회신인지 후보를 제시한다. 저장은 하지 않는다 (FR-104)."""
    try:
        parsed, _, _ = services.resolve_input(body)
    except services.InputError as error:
        raise _http_error(error) from error
    return services.match_reply(session, parsed.snapshot)


@app.post("/api/cases/{case_id}/replies", response_model=CaseDetail)
def merge_reply(case_id: str, body: MailInput, session: Session = Depends(get_session)):
    """[회신 병합] 담당자가 확인한 케이스에 회신을 합치고 재추출·재검증을 실행한다 (FR-104·FR-207)."""
    try:
        parsed, raw, source = services.resolve_input(body)
        case = services.merge_reply(session, case_id, parsed, raw, source, body.actor)
    except (services.InputError, services.NotFoundError, services.ConflictError) as error:
        raise _http_error(error) from error
    return services.case_detail(case)
