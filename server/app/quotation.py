"""견적서 생성: 요율 마스터로 금액을 계산하고 발주 측 양식(견적서_양식.xlsx)에 채워 저장한다.

- 계산은 룰 베이스만 (LLM 없음). 단가 × 수량 = 금액, 통화별 합계 (양식 'TOTAL (USD+KRW)').
- 요율이 없는 구간은 견적을 만들지 않는다 (FR-501) → None.
- 파생 필드(quoteCurrency·transitTime·freeTimeDet/Dem·validityDays·exchangeRateAsOf)를 lane_rules에서 채운다 (FR-507·508).
- 파일은 storage/<케이스ID>/quotes/<케이스ID>_견적서.xlsx → A트랙 [견적서 송부 초안]이 그대로 첨부한다.
  다시 계산하면 이전 파일은 quote-history/로 옮긴다 (송부 초안에 옛 버전이 같이 붙지 않도록).
- 케이스 ID를 Subject 줄에 표기한다 (FR-102: 화면·메일 제목·견적서 동일 표기).

PDF는 LibreOffice(있을 때만)로 같은 XLSX를 변환해 옆에 둔다 → 송부 초안에 PDF·XLSX가 함께 첨부된다 (MUST-SHIP ④).
"""

import io
import logging
import math
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import timedelta

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
from sqlalchemy import func, or_, select
from sqlalchemy.orm import Session

from . import drafts
from .config import settings
from .extraction import ports
from .models import Case, ContainerType, LaneRule, Quote, QuoteItem, QuoteTotal, Rate

# 해상운임이 견적에 반드시 들어가야 하는 조건. FOB는 해상운임을 매수인이 부담해 국내 비용만 견적한다.
FREIGHT_REQUIRED = ("CIF", "DDP", "EXW")

ITEM_FIRST_ROW, ITEM_LAST_ROW, TOTAL_ROW = 23, 34, 35  # 양식의 비용 12칸 + 합계 행
REMARK_FIRST_ROW, REMARK_LAST_ROW = 39, 47
MONEY_FORMAT = {"KRW": '"₩"#,##0', "USD": '"US$"#,##0.00'}


def _applies(column, value):
    """요율 행의 조건 칸이 비어 있으면 모든 값에 적용."""
    return or_(column.is_(None), column == value)


def find_rates(session: Session, qi, on_date) -> list[Rate]:
    rows, _reason = rates_with_reason(session, qi, on_date)
    return rows


def rates_with_reason(session: Session, qi, on_date) -> tuple[list[Rate], str | None]:
    """견적에 쓸 요율 행과, 못 쓰면 그 이유 (FR-501: 요율이 모자라면 견적을 만들지 않는다)."""
    lane = f"{qi.pol}→{qi.pod} {qi.containerType}"
    rows = session.scalars(
        select(Rate)
        .where(
            _applies(Rate.pol, qi.pol),
            _applies(Rate.pod, qi.pod),
            _applies(Rate.container_type, qi.containerType),
            _applies(Rate.incoterms, qi.incoterms),
            or_(Rate.valid_from.is_(None), Rate.valid_from <= on_date),
            or_(Rate.valid_until.is_(None), Rate.valid_until >= on_date),
        )
        .order_by(Rate.sort_no, Rate.id)
    ).all()
    # 구간(pol 또는 pod)이 지정된 요율이 하나도 없으면 '요율 없는 구간'으로 본다 — 공통 부대비용만으로 견적 금지
    if not any(r.pol or r.pod for r in rows):
        return [], f"요율 미등록 구간 {lane} (FR-501)"
    # CIF 등인데 이 도착항의 해상운임 요율이 없으면 견적 불가 (출발지 부대비용만으로 만들지 않는다)
    if qi.incoterms in FREIGHT_REQUIRED and not any(r.charge_code == "OCEAN_FREIGHT" for r in rows):
        return [], f"해상운임 요율 없음 {lane} {qi.incoterms} (FR-501)"
    # 출발지 THC가 없으면 국내 부대비용 요율이 없는 것 — 해상운임만 있는 불완전한 견적을 만들지 않는다
    if not any(r.charge_code == "THC" for r in rows):
        return [], f"국내 부대비용(THC 등) 요율 없음 {qi.pol} {qi.containerType} (FR-501)"
    return list(rows), None


def lane_rule(session: Session, qi) -> LaneRule | None:
    rows = session.scalars(
        select(LaneRule).where(
            _applies(LaneRule.pol, qi.pol), _applies(LaneRule.pod, qi.pod),
            _applies(LaneRule.container_type, qi.containerType),
        )
    ).all()
    # 조건이 많이 맞는(구체적인) 룰 우선
    return max(rows, key=lambda r: sum(x is not None for x in (r.pol, r.pod, r.container_type)), default=None)


def revenue_ton(qi) -> float:
    """LCL 청구 단위 R/T = max(CBM, 중량톤), 최소 1 R/T (예시 견적서 '1CBM 기준')."""
    cbm = qi.totalCbm
    if cbm is None and None not in (qi.boxL, qi.boxW, qi.boxH, qi.qty):
        cbm = qi.boxL * qi.boxW * qi.boxH / 1e9 * qi.qty
    ton = (qi.grossWeightKg or 0) / 1000
    return round(max(1.0, cbm or 0, ton), 3)


USABLE_RATIO = 0.88  # 컨테이너 내부 부피 중 실제 적재 가능한 비율 (산정용 보수 기준)


def container_count(session: Session, qi, stated: int | None) -> tuple[int | None, str | None]:
    """FCL 컨테이너 대수 → (대수, 산정 근거). 메일에 적힌 대수가 우선, 없으면 부피·중량으로 필요한 대수를 계산.

    6장에 대수 필드가 없어 추출 결과의 보조값(containerCount)으로 받는다. LCL이면 (None, None).
    """
    if qi.containerType in (None, "LCL"):
        return None, None
    if stated:
        return stated, None
    ct = session.get(ContainerType, qi.containerType)
    cbm = qi.totalCbm
    if cbm is None and None not in (qi.boxL, qi.boxW, qi.boxH, qi.qty):
        cbm = qi.boxL * qi.boxW * qi.boxH / 1e9 * qi.qty
    needs = [1]
    if ct and ct.max_cbm and cbm:
        needs.append(math.ceil(cbm / (ct.max_cbm * USABLE_RATIO)))
    if ct and ct.max_gross_kg and qi.grossWeightKg:
        needs.append(math.ceil(qi.grossWeightKg / ct.max_gross_kg))
    count = max(needs)
    basis = []
    if cbm:
        basis.append(f"{_fmt_num(round(cbm, 2))}CBM")
    if qi.grossWeightKg:
        basis.append(f"{_fmt_num(qi.grossWeightKg)}KG")
    note = f"컨테이너 대수: 화물 {'·'.join(basis) or '정보'} 기준 {qi.containerType} {count}대로 산정" if basis else None
    return count, note


def quantity(basis: str, qi, count: int | None = None) -> float | None:
    if basis == "PER_RT":
        return revenue_ton(qi)
    if basis == "AT_COST":
        return None
    if basis == "PER_CNTR":
        return float(count or 1)
    return 1.0  # PER_BL·PER_TRIP


def money(value: float, currency: str) -> float:
    return float(round(value)) if currency == "KRW" else round(value + 1e-9, 2)


def build(session: Session, case: Case, extraction: dict | None = None) -> dict | None:
    qi = case.quote_input
    stated = ((extraction or {}).get("fields") or {}).get("containerCount") or {}
    count, count_note = container_count(session, qi, stated.get("value") if stated.get("status") == "filled" else None)
    from .validation import today_kst  # 순환 import 방지

    today = today_kst()
    rates, _reason = rates_with_reason(session, qi, today)
    if not rates:
        return None

    rule = lane_rule(session, qi)
    qi.quoteCurrency = rule.quote_currency if rule else "USD"
    qi.transitTime = rule.transit_days if rule else None
    qi.freeTimeDet = rule.free_time_det if rule else None
    qi.freeTimeDem = rule.free_time_dem if rule else None
    qi.validityDays = rule.validity_days if rule else 14
    qi.exchangeRateAsOf = today.isoformat()  # 실제 환율은 출항일 기준 적용 — 견적서 비고에 고지 (FR-508)

    version = (session.scalar(select(func.max(Quote.version_no)).where(Quote.case_pk == case.id)) or 0) + 1
    quote = Quote(case=case, version_no=version, valid_until=(today + timedelta(days=qi.validityDays)).isoformat())
    totals: dict[str, float] = {}
    for rate in rates:
        qty = quantity(rate.basis, qi, count)
        amount = None
        if qty is not None and rate.unit_price is not None:
            amount = money(rate.unit_price * qty, rate.currency)
            totals[rate.currency] = money(totals.get(rate.currency, 0) + amount, rate.currency)
        quote.items.append(QuoteItem(
            charge_code=rate.charge_code, charge_label=rate.charge_label, basis=rate.basis,
            rate=rate.unit_price, qty=qty, amount=amount, currency=rate.currency, remark=rate.remark,
        ))
    for currency in sorted(totals, key=lambda c: (c != "USD", c)):
        quote.totals.append(QuoteTotal(currency=currency, total=totals[currency]))
    session.add(quote)
    session.flush()

    xlsx = render_xlsx(case, quote, count, count_note)
    quote.xlsx_path = _save(case, quote, "xlsx", xlsx)
    pdf_path, pdf_note = None, None
    if settings.quote_pdf:
        pdf, pdf_note = to_pdf(xlsx)
        if pdf:
            pdf_path = _save(case, quote, "pdf", pdf)
    session.flush()
    return {
        "version": quote.version_no,
        "file": quote.xlsx_path,
        "pdf": pdf_path,
        "pdfNote": pdf_note,
        "validUntil": quote.valid_until,
        "totals": {t.currency: t.total for t in quote.totals},
        "items": len(quote.items),
        "containerCount": count,
        "containerCountNote": count_note,
        "rateSources": sorted({r.source for r in rates if r.source}),
    }


# ------------------------------------------------------------------ XLSX 양식 채우기


def port_label(code: str | None) -> str:
    names = ports().get(code or "", ())
    english = [n for n in names if n.isascii() and len(n) > 2]
    return english[0] if english else (code or "")


def _fmt_num(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.2f}"


def total_text(quote: Quote) -> str:
    parts = [f"{t.currency}{t.total:,.2f}" if t.currency == "USD" else f"{t.currency}{t.total:,.0f}" for t in quote.totals]
    return " + ".join(parts)


COUNTRY = {"KR": "KOREA"}  # FOB 예시의 'BUSAN, KOREA' 표기용 (수출 견적은 선적항이 국내)


def _korean_date(iso: str | None) -> str:
    """'2026-09-30' → '2026년 9월 30일' (예시 견적서 VALIDITY 표기)"""
    if not iso:
        return ""
    y, m, d = iso.split("-")
    return f"{int(y)}년 {int(m)}월 {int(d)}일"


def _box(ws, row: int, first: str, last: str) -> None:
    thin, thick = Side(style="thin"), Side(style="medium")
    cols = [chr(c) for c in range(ord(first), ord(last) + 1)]
    for col in cols:
        ws[f"{col}{row}"].border = Border(
            left=thin if col == first else None, right=thick if col == last else None, top=thin, bottom=thick,
        )


def render_xlsx(case: Case, quote: Quote, count: int | None = None, count_note: str | None = None) -> bytes:
    """발주 측 양식(견적서_양식.xlsx)에 채운다. 머리글·표 제목은 예시 견적서(LCL·FOB PDF)를 따른다."""
    qi = case.quote_input
    items = quote.items
    if len(items) > ITEM_LAST_ROW - ITEM_FIRST_ROW + 1:
        raise ValueError(f"비용 항목이 양식 칸(12개)보다 많습니다: {len(items)}개")

    wb = load_workbook(settings.quote_template)
    ws = wb.active
    is_lcl = qi.containerType == "LCL"
    is_fob_fcl = qi.incoterms == "FOB" and not is_lcl  # FOB 예시: 국내 비용만, POL만 표기

    # 인쇄: 양식의 인쇄 범위(A:F)를 한 페이지 폭에 맞춘다 — 없으면 회사 주소·DATE·REMARK(F열)가 2쪽으로 밀린다
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    ws["B9"] = f"수 신 : {qi.customerName or ''} / {qi.contactName or ''} 님".replace(" /  님", " 님")
    ws["B10"] = "발 신 : 레오나 해운항공㈜ / 견적 담당 드림"
    ws["B12"] = f"Subject : 해상 수출 운임 제안서  [{case.case_id}]"
    _box(ws, 12, "B", "D")  # 양식의 Subject 상자는 도형이라 저장 시 빠진다(openpyxl) → 셀 테두리로 대신 그린다

    size = f" / {qi.boxL} X {qi.boxW} X {qi.boxH}MM" if None not in (qi.boxL, qi.boxW, qi.boxH) else ""
    cbm = f" / {_fmt_num(qi.totalCbm)}CBM" if qi.totalCbm is not None else ""
    cargo = f": {qi.qty} {qi.qtyUnit}{size} / {_fmt_num(qi.grossWeightKg)}KGS{cbm}"
    if is_fob_fcl:
        country = COUNTRY.get((qi.pol or "")[:2])
        ws["B17"] = "** POL"
        ws["C17"] = f": {port_label(qi.pol)}{', ' + country if country else ''}"
        ws["C18"] = ": BY CONTAINER SHIPMENT"
        ws["B19"] = "** REQUIRED CNTR"
        ws["C19"] = f": {qi.containerType}' X {count or 1}"
        headers = ("CONTAINER", "Q/T", "TOTAL")
    else:
        ws["C17"] = f": {port_label(qi.pol)} - {port_label(qi.pod)}"
        if is_lcl:
            ws["C18"] = ": BY CONTAINER SHIPMENT (LCL)"
            rt = next((i.qty for i in items if i.basis == "PER_RT" and i.qty), None)
            ws["C19"] = cargo + (f" ({_fmt_num(rt)} R/T 기준)" if rt else "")
            headers = ("RATE", "R/T", "AMOUNT")
        else:
            ws["C18"] = f": BY CONTAINER SHIPMENT ({qi.containerType} X {count or 1})"
            ws["C19"] = cargo
            headers = ("RATE", "Q'TY", "AMOUNT")
    ws["C20"] = f": {qi.incoterms}"  # 발주 측 양식은 이 줄(PAYMENT TERM)에 인코텀즈를 적는다 (예시 견적서 기준)
    ws["F21"] = f"DATE : {quote.created_at:%Y. %m. %d}" if quote.created_at else "DATE :"
    ws["E21"] = None
    ws["C22"], ws["D22"], ws["E22"] = headers

    for row in range(ITEM_FIRST_ROW, ITEM_LAST_ROW + 1):
        for col in "BCDEF":
            ws[f"{col}{row}"] = None
    for row, item in zip(range(ITEM_FIRST_ROW, ITEM_LAST_ROW + 1), items):
        ws[f"B{row}"] = item.charge_label
        if item.basis == "AT_COST" or item.rate is None:
            ws.merge_cells(f"C{row}:E{row}")  # FOB 예시처럼 단가~금액 칸을 합쳐 AT COST
            ws[f"C{row}"] = "AT COST"
            ws[f"C{row}"].alignment = Alignment(horizontal="center", vertical="center")
        else:
            ws[f"C{row}"] = item.rate
            ws[f"C{row}"].number_format = MONEY_FORMAT.get(item.currency, "#,##0.00")
            ws[f"D{row}"] = item.qty
            ws[f"E{row}"] = item.amount
            ws[f"E{row}"].number_format = MONEY_FORMAT.get(item.currency, "#,##0.00")
        ws[f"F{row}"] = item.remark or {"PER_RT": "PER R/T", "PER_CNTR": "PER CNTR", "PER_BL": "PER B/L",
                                         "PER_TRIP": "PER TRIP"}.get(item.basis, "")
    for row in range(ITEM_FIRST_ROW + len(items), ITEM_LAST_ROW + 1):
        ws.row_dimensions[row].hidden = True  # 쓰지 않은 비용 줄은 숨긴다 (예시처럼 쓴 줄만 보이게)
    ws[f"C{TOTAL_ROW}"] = total_text(quote)

    remarks = ["* 실제 출항일 환율 적용", "* 화물 DETAIL의 변경에 따라 상기 견적이 달라 질 수 있습니다."]
    if qi.transitTime:
        remarks.append(f"* T/T : 약 {qi.transitTime}일")
    if qi.freeTimeDet or qi.freeTimeDem:
        remarks.append(f"* FREE TIME : DET {qi.freeTimeDet or '-'}일 / DEM {qi.freeTimeDem or '-'}일")
    if count_note:
        remarks.append(f"* {count_note}")
    remarks.append(f"* VALIDITY : {_korean_date(quote.valid_until)}")
    remarks.append(f"* 견적번호 : {case.case_id} (v{quote.version_no})")
    for offset, row in enumerate(range(REMARK_FIRST_ROW, REMARK_LAST_ROW + 1)):
        ws[f"B{row}"] = remarks[offset] if offset < len(remarks) else None

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _save(case: Case, quote: Quote, ext: str, data: bytes) -> str:
    """storage/<케이스ID>/quotes/<케이스ID>_견적서.<ext> — 이전 버전은 quote-history/로 옮긴다."""
    filename = f"{case.case_id}_견적서.{ext}"
    current = settings.storage_dir / case.case_id / "quotes" / filename
    if current.exists():
        history = settings.storage_dir / case.case_id / "quote-history"
        history.mkdir(parents=True, exist_ok=True)
        shutil.move(current, history / f"{case.case_id}_견적서_v{quote.version_no - 1}.{ext}")
    return drafts.save_quote_file(case, filename, data)


# ------------------------------------------------------------------ PDF (LibreOffice)

log = logging.getLogger(__name__)
_WINDOWS_SOFFICE = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


def find_soffice() -> str | None:
    if settings.soffice_path:
        return settings.soffice_path if Path(settings.soffice_path).exists() else None
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    return next((p for p in _WINDOWS_SOFFICE if os.path.exists(p)), None)


def to_pdf(xlsx: bytes) -> tuple[bytes | None, str | None]:
    """양식 XLSX → PDF. LibreOffice가 없거나 실패하면 (None, 사유) — 견적서 XLSX는 그대로 쓴다."""
    soffice = find_soffice()
    if soffice is None:
        return None, "LibreOffice가 없어 PDF를 만들지 않았습니다 (XLSX만 생성)"
    with tempfile.TemporaryDirectory(prefix="leona-pdf-") as tmp:
        src = Path(tmp) / "quote.xlsx"
        src.write_bytes(xlsx)
        profile = (Path(tmp) / "profile").as_uri()  # 담당자가 LibreOffice를 켜 둬도 충돌하지 않게 별도 프로필
        try:
            subprocess.run(
                [soffice, "--headless", "--norestore", f"-env:UserInstallation={profile}",
                 "--convert-to", "pdf", "--outdir", tmp, str(src)],
                capture_output=True, timeout=settings.pdf_timeout_s, check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            log.warning("pdf conversion failed: %s", error)
            return None, f"PDF 변환 실패 ({error.__class__.__name__}) — XLSX만 생성"
        out = Path(tmp) / "quote.pdf"
        if not out.exists():
            return None, "PDF 변환 결과가 없습니다 — XLSX만 생성"
        return out.read_bytes(), None
