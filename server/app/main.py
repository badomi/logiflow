"""A트랙 백엔드 API 서버.

실행: server 폴더에서  .venv\\Scripts\\python -m uvicorn app.main:app --reload --port 8000
API 문서(자동 생성): http://localhost:8000/docs
"""

from contextlib import asynccontextmanager

from fastapi import Depends, FastAPI, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from . import drafts, services
from .db import get_session, init_db
from .schemas import (
    ActorRequest,
    CaseDetail,
    CaseSummary,
    CreateCaseResult,
    DraftOut,
    MailEventRequest,
    MailInput,
    ReplyMatchResult,
    SupplementRequest,
)


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


@app.post("/api/cases/{case_id}/drafts/supplement", response_model=DraftOut)
def supplement_draft(case_id: str, body: SupplementRequest, session: Session = Depends(get_session)):
    """[보완 요청 초안] 제목에 케이스 ID, 누락 문항만 번호로 나열, 국·영문 병기 (FR-304·305·510)."""
    try:
        case = services.get_case(session, case_id)
        return drafts.supplement_draft(session, case, body.questions, body.actor)
    except (services.NotFoundError, drafts.DraftError) as error:
        raise _http_error(error) from error


@app.post("/api/cases/{case_id}/drafts/quote", response_model=DraftOut)
def quote_draft(case_id: str, body: ActorRequest, session: Session = Depends(get_session)):
    """[견적서 송부 초안] 견적서 PDF·XLSX 첨부 (FR-505)."""
    try:
        case = services.get_case(session, case_id)
        return drafts.quote_draft(session, case, body.actor)
    except (services.NotFoundError, drafts.DraftError) as error:
        raise _http_error(error) from error


@app.post("/api/cases/{case_id}/mail-events", response_model=CaseDetail)
def mail_event(case_id: str, body: MailEventRequest, session: Session = Depends(get_session)):
    """[초안 저장·발송 기록] Outlook 작성 창과 보내기 이벤트가 호출한다. 기록만 하고 발송하지 않는다 (FR-505)."""
    try:
        case = services.get_case(session, case_id)
        drafts.record_mail_event(session, case, body.event_type, body.kind, body.subject, body.actor)
    except (services.NotFoundError, drafts.DraftError) as error:
        raise _http_error(error) from error
    return services.case_detail(case)


@app.post("/api/cases/{case_id}/quotes")
async def upload_quote(case_id: str, file: UploadFile, session: Session = Depends(get_session)):
    """견적서 파일 등록. C트랙 견적서 생성(FR-504) 결과를 케이스에 붙이는 입구 (연결 전에는 테스트용)."""
    try:
        case = services.get_case(session, case_id)
        path = drafts.save_quote_file(case, file.filename or "quote.pdf", await file.read())
    except (services.NotFoundError, drafts.DraftError) as error:
        raise _http_error(error) from error
    return {"path": path}


@app.get("/api/cases/{case_id}/quotes/{filename}")
def download_quote(case_id: str, filename: str, session: Session = Depends(get_session)):
    """패널이 초안에 첨부할 견적서를 내려받는다. 케이스의 견적서 폴더에 있는 파일만 준다."""
    try:
        case = services.get_case(session, case_id)
    except services.NotFoundError as error:
        raise _http_error(error) from error
    match = next((f for f in drafts.quote_files(case) if f.name == filename), None)
    if match is None:
        raise HTTPException(status_code=404, detail="견적서 파일이 없습니다.")
    return FileResponse(match, filename=match.name)


@app.post("/api/cases/{case_id}/replies", response_model=CaseDetail)
def merge_reply(case_id: str, body: MailInput, session: Session = Depends(get_session)):
    """[회신 병합] 담당자가 확인한 케이스에 회신을 합치고 재추출·재검증을 실행한다 (FR-104·FR-207)."""
    try:
        parsed, raw, source = services.resolve_input(body)
        case = services.merge_reply(session, case_id, parsed, raw, source, body.actor)
    except (services.InputError, services.NotFoundError, services.ConflictError) as error:
        raise _http_error(error) from error
    return services.case_detail(case)
