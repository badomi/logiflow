"""제공 요율표 → 새 DB → C트랙의 병합 경계 검증."""
import io
from datetime import date

import pytest
from openpyxl import load_workbook
from sqlalchemy import select

from app import rates, validation
from app.db import SessionLocal
from app.models import Rate
from app.sample_rate_quote import SOURCE
from .test_outlook_sample_quote import BODY, import_client_rates
from .test_pipeline_flow import connect, create, text_mail


def test_client_rate_export_preserves_new_db_columns(client, connect):
    import_client_rates(client)
    with SessionLocal() as session:
        def values():
            return {r.source_ref: (r.basis, r.unit_price, *[getattr(r, k) for k in rates.RATE_METADATA])
                    for r in session.scalars(select(Rate))}
        before = values()
        data = rates.export(session)
        rates.import_workbook(session, data)
        session.expire_all()
        assert values() == before


@pytest.mark.parametrize("cell,value", [("O8", -1), ("P8", 1.5), ("I8", -430)])
def test_invalid_import_is_atomic(client, connect, cell, value):
    import_client_rates(client)
    wb = load_workbook(SOURCE)
    wb["OceanFreight"][cell] = value
    buf = io.BytesIO(); wb.save(buf)
    response = client.post("/api/rates/import", files={"file": ("rates.xlsx", buf.getvalue())})
    assert response.status_code == 400
    case = create(client, text_mail(BODY))
    assert case["extraction"]["validation"]["quote"]["totals"] == {"USD": 870}


def test_db_schedule_rejects_departure_before_ready(client, connect):
    import_client_rates(client)
    case = create(client, text_mail(BODY.replace("2026-10-05", "2026-10-08")))
    assert case["status"] == "보류"
    assert "스케줄" in case["extraction"]["validation"]["hold"][0]["message"]


def test_db_schedule_rejects_inconsistent_eta(client, connect):
    import_client_rates(client)
    with SessionLocal() as session:
        row = session.scalar(select(Rate).where(Rate.carrier == "SITC", Rate.valid_from == date(2026,10,1)))
        row.eta = "2026-10-12"
        session.commit()
    case = create(client, text_mail(BODY))
    assert case["status"] == "보류"
    assert "ETA" in case["extraction"]["validation"]["hold"][0]["message"]


def test_validity_capped_by_source_and_pdf_failure_keeps_no_old_attachment(client, connect, monkeypatch):
    import_client_rates(client)
    monkeypatch.setattr(validation, "today_kst", lambda: date(2026,10,2))
    case = create(client, text_mail(BODY))
    quote = case["extraction"]["validation"]["quote"]
    assert quote["validUntil"] == "2026-10-15"
    assert quote["etd"] == quote["schedule"]["etd"] == "2026-10-07"
    assert all(line["source"] for line in quote["lines"])
