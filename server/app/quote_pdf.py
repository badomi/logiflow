"""견적서 PDF 내장 엔진 — LibreOffice가 없거나 변환에 실패할 때 쓴다 (MUST-SHIP ④: PDF·XLSX).

quotation.render_xlsx()가 채운 양식 XLSX를 그대로 읽어 같은 배치로 그린다.
→ 문구·비용 행·비고는 XLSX와 항상 같다 (PDF용 내용을 따로 만들지 않는다).

한글 폰트는 PDF에 내장한다 (내장하지 않으면 고객 PC·모바일 뷰어에서 한글이 깨질 수 있음).
  Windows: 맑은 고딕(malgun.ttf / malgunbd.ttf) 자동 사용
  그 밖: 환경 변수 LEONA_PDF_FONT=<TTF 경로> 또는 나눔고딕 설치
"""

from __future__ import annotations

import io
import logging
import os
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

log = logging.getLogger(__name__)

# (본문 TTF, 굵은 TTF) 후보 — 앞에서부터 있는 것을 쓴다
_FONTS = [
    (os.environ.get("LEONA_PDF_FONT", ""), os.environ.get("LEONA_PDF_FONT_BOLD", "")),
    ("C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/malgunbd.ttf"),
    ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
    ("/Library/Fonts/NanumGothic.ttf", "/Library/Fonts/NanumGothicBold.ttf"),
    ("/System/Library/Fonts/AppleSDGothicNeo.ttc", ""),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", ""),
]
_font_name: str | None = None


def _font() -> str:
    """한글 폰트 등록 (처음 한 번). 굵은 글꼴이 없으면 본문 글꼴로 대신한다."""
    global _font_name
    if _font_name:
        return _font_name
    for regular, bold in _FONTS:
        if not regular or not Path(regular).exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont("LeonaKR", regular, subfontIndex=0))
        except Exception:  # noqa: BLE001 — 읽을 수 없는 폰트면 다음 후보
            continue
        bold_name = "LeonaKR"
        if bold and Path(bold).exists():
            try:
                pdfmetrics.registerFont(TTFont("LeonaKR-Bold", bold, subfontIndex=0))
                bold_name = "LeonaKR-Bold"
            except Exception:  # noqa: BLE001
                pass
        registerFontFamily("LeonaKR", normal="LeonaKR", bold=bold_name, italic="LeonaKR", boldItalic=bold_name)
        _font_name = "LeonaKR"
        return _font_name
    log.warning("내장 가능한 한글 TTF 없음 → CID 폰트 사용 (일부 뷰어에서 한글이 깨질 수 있음). LEONA_PDF_FONT 설정 권장")
    pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
    _font_name = "HYGothic-Medium"
    return _font_name


# ------------------------------------------------------------------ 셀 값 → 표시 문자열

def _fmt(cell) -> str:
    """양식의 숫자 서식("₩"#,##0 / "US$"#,##0.00)을 흉내 내 문자열로."""
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        fmt = cell.number_format or ""
        prefix = "₩" if "₩" in fmt else "US$" if "US$" in fmt else "€" if "€" in fmt else ""
        if prefix or "#,##0" in fmt:
            decimals = 2 if "0.00" in fmt else 0
            return f"{prefix}{v:,.{decimals}f}"
        return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.3f}".rstrip("0").rstrip(".")
    if hasattr(v, "strftime"):
        return v.strftime("%Y. %m. %d")
    return str(v)


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ------------------------------------------------------------------ 본체

def render_pdf(xlsx: bytes, item_rows: tuple[int, int] = (23, 34), total_row: int = 35,
               remark_rows: tuple[int, int] = (38, 49)) -> bytes:
    """채워진 양식 XLSX(bytes) → PDF(bytes). 행 번호는 quotation.py의 양식 상수와 같다."""
    font = _font()
    wb = load_workbook(io.BytesIO(xlsx))
    ws = wb.active

    def val(ref: str) -> str:
        return _fmt(ws[ref]).strip()

    def hidden(row: int) -> bool:
        return bool(ws.row_dimensions[row].hidden)

    merged_ce = {m.min_row for m in ws.merged_cells.ranges
                 if m.min_col == 3 and m.max_col == 5 and m.min_row == m.max_row}

    st = ParagraphStyle("t", fontName=font, fontSize=8.5, leading=12)
    st_r = ParagraphStyle("r", parent=st, alignment=2)
    st_c = ParagraphStyle("c", parent=st, alignment=1, leading=10)
    b = lambda t: f"<b>{_esc(t)}</b>"  # noqa: E731

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                            topMargin=18 * mm, bottomMargin=14 * mm, title=val("B12").replace("Subject : ", ""))
    width = A4[0] - 32 * mm
    story: list = []

    # 머리글: 로고 + 회사 주소 (F2~F6)
    logo = None
    if ws._images:
        try:
            data = ws._images[0]._data()
            logo = Image(io.BytesIO(data), width=58 * mm, height=23 * mm)
        except Exception:  # noqa: BLE001 — 로고를 못 읽어도 PDF는 만든다
            log.warning("양식 로고를 읽지 못해 생략")
    company = val("F2")
    address = "<br/>".join(_esc(val(f"F{r}")) for r in (4, 5, 6) if val(f"F{r}"))
    head = Table([[logo or "", Paragraph(f"<i>{b(company)}</i><br/><br/>{address}", st_r)]],
                 colWidths=[width * 0.5, width * 0.5])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [head, Spacer(1, 9 * mm)]

    # 수신·발신·Subject·인사말
    for ref in ("B9", "B10"):
        if val(ref):
            story.append(Paragraph(b(val(ref)), st))
    story.append(Spacer(1, 3 * mm))
    subject = Table([[Paragraph(b(val("B12")), st)]], colWidths=[width * 0.6], hAlign="LEFT")
    subject.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                                 ("LINEBELOW", (0, 0), (-1, -1), 1.6, colors.black),
                                 ("LINEAFTER", (0, 0), (-1, -1), 1.6, colors.black)]))
    story += [subject, Spacer(1, 3 * mm)]
    for ref in ("B14", "B15"):
        if val(ref):
            story.append(Paragraph(_esc(val(ref)), st))
    story.append(Spacer(1, 4 * mm))

    # ** POL / SHIPPING METHOD / CARGO DETAIL / PAYMENT TERM (B17~C20) + DATE (E21·F21)
    info = [[Paragraph(_esc(val(f"B{r}")), st), Paragraph(_esc(val(f"C{r}")), st)]
            for r in range(17, 21) if not hidden(r) and (val(f"B{r}") or val(f"C{r}"))]
    date = " ".join(x for x in (val("E21"), val("F21")) if x)
    info.append(["", Paragraph(b(date), st_r)])
    t = Table(info, colWidths=[width * 0.27, width * 0.73])
    t.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    story += [t, Spacer(1, 1 * mm)]

    # 비용 표: 머리글(22행) + 쓴 비용 줄 + TOTAL
    header_row = item_rows[0] - 1
    rows = [[Paragraph(b(val(f"{c}{header_row}")), st_c) for c in "BCDEF"]]
    spans = []
    for r in range(item_rows[0], item_rows[1] + 1):
        if hidden(r) or not val(f"B{r}"):
            continue
        i = len(rows)
        if r in merged_ce:
            rows.append([Paragraph(b(val(f"B{r}")), st_c), Paragraph(_esc(val(f"C{r}")), st_c), "", "",
                         Paragraph(_esc(val(f"F{r}")), st_c)])
            spans.append(("SPAN", (1, i), (3, i)))
        else:
            rows.append([Paragraph(b(val(f"B{r}")), st_c)]
                        + [Paragraph(_esc(val(f"{c}{r}")), st_c) for c in "CDEF"])
    n = len(rows)
    rows.append([Paragraph(b(val(f"B{total_row}")), st_c), Paragraph(b(val(f"C{total_row}")), st_c), "", "",
                 Paragraph(b(val(f"F{total_row}")), st_c)])
    table = Table(rows, colWidths=[width * 0.25, width * 0.18, width * 0.09, width * 0.18, width * 0.30],
                  repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
        ("BOX", (0, 0), (-1, -1), 1.2, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#BFBFBF")),
        ("BACKGROUND", (0, n), (-1, n), colors.HexColor("#FFFF00")),
        ("SPAN", (1, n), (3, n)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        *spans,
    ]))
    story += [table, Spacer(1, 9 * mm)]

    # Remark ~ 감사합니다 ~ 직인명판 (38~49행)
    first, last = remark_rows
    for r in range(first, last + 1):
        if hidden(r):
            continue
        left, right = val(f"B{r}"), val(f"F{r}")
        if r == last - 1 and left:
            story.append(Spacer(1, 4 * mm))          # '* 감사합니다.' 앞 여백 (양식처럼)
        if r == first and left:
            story += [Paragraph(f"<u>{b(left)}</u>", st), Spacer(1, 1 * mm)]
        elif left:
            story.append(Paragraph(_esc(left), st))
        if right:
            story += [Spacer(1, 3 * mm), Paragraph(_esc(right), st_r)]
    doc.build(story)
    return buf.getvalue()
