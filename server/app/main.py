"""A트랙 백엔드 API 서버.

실행: server 폴더에서  .venv\\Scripts\\python -m uvicorn app.main:app --reload --port 8000
API 문서(자동 생성): http://localhost:8000/docs
"""

from contextlib import asynccontextmanager
import json

from fastapi import BackgroundTasks, Depends, FastAPI, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from . import drafts, pipeline, services
from .config import settings
from .db import SessionLocal, get_session, init_db
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
    QuoteValidityRequest,
)


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    connect_pipeline()
    yield


def connect_pipeline() -> None:
    """컨테이너 마스터 등록(6장 containerType 제약) + LLM_ENABLED=true면 추출·검증·견적 연결."""
    from .seed import seed_containers

    with SessionLocal() as session:
        seed_containers(session)
        session.commit()
    if settings.extraction_mode in ("rules", "hybrid"):
        from .extraction import HybridExtractor
        from .validation import RuleValidator

        pipeline.register(extractor=HybridExtractor(llm_mode=settings.extraction_mode), validator=RuleValidator())


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


@app.get("/api/health/detail")
def health_detail(session: Session = Depends(get_session)) -> dict:
    """설정 점검: 추출 모드·LLM 연결·모델 설치·요율·PDF 도구·.env — 문제가 있으면 warnings에 할 일을 적는다."""
    from .diagnostics import collect

    return collect(session)


@app.post("/api/cases", response_model=CreateCaseResult)
def create_case(body: MailInput, background: BackgroundTasks, session: Session = Depends(get_session)):
    """[케이스 생성] 담당자가 패널에서 누른다 (FR-101). ID 발급(FR-102)·원문 보관(FR-103)·중복 방지(FR-105).

    케이스는 즉시 만들어 응답하고, 항목 추출은 응답 뒤 백그라운드에서 실행한다.
    패널은 GET /api/cases/{id}의 extraction 상태로 진행·결과·걸린 시간(60초 기준)을 확인한다.
    """
    try:
        parsed, raw, source = services.resolve_input(body)
        case, duplicate, job = services.create_case(session, parsed, raw, source, body.actor)
    except (services.InputError, services.NotFoundError, services.ConflictError) as error:
        raise _http_error(error) from error
    if job:
        background.add_task(pipeline.execute, job)
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


@app.get("/api/cases/{case_id}/mails/{index}/raw")
def download_raw_mail(case_id: str, index: int, session: Session = Depends(get_session)):
    """[원문 다시 열기] 케이스의 메일 원문(.eml). Outlook 등 메일 프로그램으로 열 수 있다 (FR-103)."""
    try:
        path = services.raw_mail_path(session, case_id, index)
    except services.NotFoundError as error:
        raise _http_error(error) from error
    return FileResponse(path, media_type="message/rfc822", filename=f"{case_id}-mail-{index:02d}.eml")


@app.get("/api/cases/{case_id}/attachments/{attachment_id}")
def download_attachment(case_id: str, attachment_id: int, session: Session = Depends(get_session)):
    """[첨부 다시 열기] 케이스에 저장된 첨부파일 (FR-103)."""
    try:
        path, filename = services.attachment_path(session, case_id, attachment_id)
    except services.NotFoundError as error:
        raise _http_error(error) from error
    return FileResponse(path, filename=filename)


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


@app.get("/api/drafts/pending", response_model=DraftOut | None)
def pending_draft(actor: str, session: Session = Depends(get_session)):
    """이 담당자가 방금 패널에서 준비한(아직 저장 기록이 없는) 초안. 작성 창 버튼의 보조 경로."""
    return drafts.latest_pending_draft(session, actor)


@app.post("/api/cases/{case_id}/mail-events", response_model=CaseDetail)
def mail_event(case_id: str, body: MailEventRequest, session: Session = Depends(get_session)):
    """[초안 저장·발송 기록] Outlook 작성 창과 보내기 이벤트가 호출한다. 기록만 하고 발송하지 않는다 (FR-505)."""
    try:
        case = services.get_case(session, case_id)
        occurred_at = body.occurred_at
        if body.eml_base64:
            # 보낸 메일 원문의 Date 헤더가 실제 발송 시각이다 (Outlook의 생성 시각은 초안을 만든 시각)
            parsed, _, _ = services.resolve_input(MailInput(eml_base64=body.eml_base64))
            occurred_at = parsed.snapshot.received_at or occurred_at
        drafts.record_mail_event(
            session,
            case,
            body.event_type,
            body.kind,
            body.subject,
            body.actor,
            internet_message_id=body.internet_message_id,
            occurred_at=occurred_at,
        )
    except (services.NotFoundError, services.InputError, drafts.DraftError) as error:
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
def merge_reply(
    case_id: str, body: MailInput, background: BackgroundTasks, session: Session = Depends(get_session)
):
    """[회신 병합] 담당자가 확인한 케이스에 회신을 합치고 재추출·재검증을 백그라운드로 실행한다 (FR-104·FR-207)."""
    try:
        parsed, raw, source = services.resolve_input(body)
        case, job = services.merge_reply(session, case_id, parsed, raw, source, body.actor)
    except (services.InputError, services.NotFoundError, services.ConflictError) as error:
        raise _http_error(error) from error
    if job:
        background.add_task(pipeline.execute, job)
    return services.case_detail(case)


# ---------------------------------------------------------------- B트랙: 추출 결과 조회·수동 보정 (FR-202·206)


class FieldCorrection(ActorRequest):
    value: str | int | float | None = None


@app.get("/api/cases/{case_id}/fields")
def get_fields(case_id: str, session: Session = Depends(get_session)) -> list[dict]:
    """필드별 현재 값·근거 원문·점수·상태 (검토 필요 값은 candidate로)."""
    from .extraction import field_view

    try:
        return field_view(session, services.get_case(session, case_id))
    except services.NotFoundError as error:
        raise _http_error(error) from error


@app.post("/api/cases/{case_id}/fields/{field}", response_model=CaseDetail)
def correct_field(case_id: str, field: str, body: FieldCorrection, background: BackgroundTasks,
                  session: Session = Depends(get_session)):
    """담당자가 필드 값을 고친다 → 수정 이력 저장 → 검증 다시 실행 (재추출이 이 값을 덮어쓰지 않음)."""
    from .extraction import correct_field as correct
    from .extraction.extractor import CorrectionError

    try:
        case = services.get_case(session, case_id)
        correct(session, case, field, body.value, body.actor)
        session.commit()
    except (services.NotFoundError, CorrectionError) as error:
        raise _http_error(error) from error
    job = pipeline.start(session, case, trigger="FIELD_CORRECTED", actor=body.actor)
    background.add_task(pipeline.execute, job)
    return services.case_detail(case)


RERUN_STUCK_AFTER_S = 120  # 이 시간이 지나도 '추출 중'이면 서버가 도중에 꺼진 것으로 보고 다시 실행을 허용


@app.post("/api/cases/{case_id}/quote-validity", response_model=CaseDetail)
def set_quote_validity(case_id: str, body: QuoteValidityRequest, background: BackgroundTasks,
                       session: Session = Depends(get_session)):
    """케이스별 유효일수 변경을 기록하고 PDF/XLSX를 새 견적 버전으로 재생성한다."""
    from .models import CaseEvent, QuoteInput
    try:
        case = services.get_case(session, case_id)
    except services.NotFoundError as error:
        raise _http_error(error) from error
    if services.extraction_state(case).state == "running":
        raise HTTPException(status_code=409, detail="현재 처리 중입니다. 완료된 뒤 유효기간을 적용해 주세요.")
    record = case.quote_input or QuoteInput(case=case)
    before = record.validityDays
    record.validityDays = body.validity_days
    session.add(record)
    session.add(CaseEvent(case=case, event_type="QUOTE_VALIDITY_CHANGED", actor=body.actor,
                          detail=json.dumps({"before": before, "after": body.validity_days})))
    job = pipeline.start(session, case, trigger="FIELD_CORRECTED", actor=body.actor)
    background.add_task(pipeline.execute, job)
    return services.case_detail(case)


@app.post("/api/cases/{case_id}/extract", response_model=CaseDetail)
def rerun_extraction(case_id: str, body: ActorRequest, background: BackgroundTasks,
                     session: Session = Depends(get_session)):
    """[다시 추출] 추출→검증을 다시 실행한다 (NFR-03 재실행). 설정 변경 뒤·LLM 시간 초과 뒤에 쓴다.

    담당자가 고친 필드(FR-206)는 다시 추출해도 유지된다.
    """
    from datetime import datetime, timezone

    try:
        case = services.get_case(session, case_id)
    except services.NotFoundError as error:
        raise _http_error(error) from error
    state = services.extraction_state(case)
    if state.state == "running" and state.started_at is not None:
        started = state.started_at if state.started_at.tzinfo else state.started_at.replace(tzinfo=timezone.utc)
        if (datetime.now(timezone.utc) - started).total_seconds() < RERUN_STUCK_AFTER_S:
            raise HTTPException(status_code=409, detail="이미 추출 중입니다. 끝난 뒤 다시 눌러 주세요.")
    job = pipeline.start(session, case, trigger="MANUAL_RERUN", actor=body.actor)
    background.add_task(pipeline.execute, job)
    return services.case_detail(case)


@app.get("/api/extraction/review-report")
def extraction_review_report(days: int = 30, session: Session = Depends(get_session)) -> dict:
    """놓친 표현 모아 보기: 찾았지만 읽지 못한 원문(unparsed)·확신 낮은 값(lowConfidence)을 항목별로"""
    from .extraction.report import collect

    return collect(session, days)


# ---------------------------------------------------------------- 요율 엑셀 관리 (NFR-06)

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _xlsx_response(data: bytes, filename: str):
    from urllib.parse import quote

    from fastapi.responses import Response

    return Response(data, media_type=XLSX_MEDIA,
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"})


@app.get("/api/rates/template")
def rates_template():
    """빈 요율 양식 (작성 안내 시트 포함)"""
    from . import rates

    return _xlsx_response(rates.template(), "요율_양식.xlsx")


@app.get("/api/rates/export")
def rates_export(session: Session = Depends(get_session)):
    """지금 DB에 있는 요율 — 고쳐서 그대로 다시 올리면 된다"""
    from . import rates

    return _xlsx_response(rates.export(session), "현재_요율.xlsx")


@app.get("/api/rates/summary")
def rates_summary(session: Session = Depends(get_session)) -> list[dict]:
    """출처별 요율 개수·적용 기간·만료 여부"""
    from . import rates

    return rates.summary(session)


@app.post("/api/rates/import")
async def rates_import(file: UploadFile, session: Session = Depends(get_session)) -> dict:
    """요율 엑셀 올리기. 틀린 줄이 하나라도 있으면 아무것도 바꾸지 않고 400 + 틀린 곳 목록."""
    from . import rates

    try:
        return rates.import_workbook(session, await file.read())
    except rates.RateImportError as error:
        raise HTTPException(status_code=400, detail={"message": str(error), "errors": error.errors}) from error
