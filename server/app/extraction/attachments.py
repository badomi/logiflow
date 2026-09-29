"""첨부 텍스트 추출 (FR-204, S): PDF·XLSX(패킹리스트 등)의 글자를 읽어 추출 대상에 넣는다.

스캔 이미지 PDF(글자 없음)는 빈 문자열 — OCR은 선택 사항이라 넣지 않았다.
표는 셀을 탭으로 이어 한 줄로 만든다 → rules.py가 '라벨<TAB>값' 줄로 읽는다.
"""

import csv
import io
import logging
from pathlib import Path

log = logging.getLogger(__name__)
MAX_CHARS = 8_000  # 60초 안에 끝나도록 첨부 1개당 길이 제한
SUPPORTED = {".pdf", ".xlsx", ".xlsm", ".csv", ".txt"}


def extract_text(filename: str, data: bytes) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED:
        return ""
    try:
        if suffix == ".pdf":
            text = _pdf(data)
        elif suffix in (".xlsx", ".xlsm"):
            text = _xlsx(data)
        elif suffix == ".csv":
            text = "\n".join("\t".join(row) for row in csv.reader(io.StringIO(_decode(data))))
        else:
            text = _decode(data)
    except Exception as error:  # 깨진 첨부 하나 때문에 추출 전체가 실패하지 않게
        log.warning("attachment text failed %s: %s", filename, error)
        return ""
    return text[:MAX_CHARS]


def _decode(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _pdf(data: bytes) -> str:
    from pypdf import PdfReader  # BSD-3

    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages[:10])


def _xlsx(data: bytes) -> str:
    from openpyxl import load_workbook  # MIT

    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    lines = []
    for ws in wb.worksheets[:5]:
        for row in ws.iter_rows(values_only=True, max_row=300):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                lines.append("\t".join(cells))
    return "\n".join(lines)
