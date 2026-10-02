"""WBS 4.9 견적서 PDF — LibreOffice가 없어도 내장 엔진으로 PDF가 만들어지는지."""

import io

from openpyxl import load_workbook

from app import quotation
from app.config import settings
from app.quote_pdf import render_pdf


def filled_template() -> bytes:
    """render_xlsx와 같은 칸에 값을 채운 양식 (DB 없이)."""
    wb = load_workbook(settings.quote_template)
    ws = wb.active
    ws["B9"] = "수 신 : 한빛무역 / 김민수 님"
    ws["B12"] = "Subject : 해상 수출 운임 제안서  [LQ-2026-1001-001]"
    ws["C17"], ws["C18"], ws["C20"] = ": BUSAN - SINGAPORE", ": BY CONTAINER SHIPMENT (40HQ X 1)", ": CIF"
    ws["F21"], ws["E21"] = "DATE : 2026. 10. 01", None
    for r in range(23, 35):
        for c in "BCDEF":
            ws[f"{c}{r}"] = None
    ws["B23"], ws["C23"], ws["D23"], ws["E23"], ws["F23"] = "OCEAN FREIGHT", 750.0, 1, 750.0, "PER CNTR"
    ws["C23"].number_format = ws["E23"].number_format = '"US$"#,##0.00'
    ws["B24"], ws["C24"], ws["D24"], ws["E24"], ws["F24"] = "THC", 410000, 1, 410000, "PER CNTR"
    ws["C24"].number_format = ws["E24"].number_format = '"₩"#,##0'
    ws["B25"], ws["F25"] = "CUSTOMS CLEARANCE FEE", "INV.V x 1/1,000"
    ws.merge_cells("C25:E25")
    ws["C25"] = "AT COST"
    for r in range(26, 35):
        ws.row_dimensions[r].hidden = True
    ws["C35"] = "USD750.00 + KRW410,000"
    ws["B39"] = "* 실제 출항일 환율 적용"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_render_pdf_from_filled_template():
    pdf = render_pdf(filled_template())
    assert pdf.startswith(b"%PDF") and len(pdf) > 5_000


def test_to_pdf_falls_back_without_libreoffice(monkeypatch):
    monkeypatch.setattr(quotation, "find_soffice", lambda: None)
    pdf, note = quotation.to_pdf(filled_template())
    assert pdf is not None and pdf.startswith(b"%PDF")
    assert "내장 엔진" in note


def test_to_pdf_keeps_xlsx_only_if_builtin_fails(monkeypatch):
    monkeypatch.setattr(quotation, "find_soffice", lambda: None)
    pdf, note = quotation.to_pdf(b"not an xlsx")
    assert pdf is None and "XLSX만 생성" in note
