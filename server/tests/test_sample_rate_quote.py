from types import SimpleNamespace

import pytest

from app.sample_rate_quote import SOURCE, calculate


def request(**changes):
    values = dict(pol="CNSHA", pod="KRPUS", containerType="20FT GP", incoterms="CIF",
                  cargoReadyDate="2026-10-05", invoiceValue=15000, ccy="USD")
    return SimpleNamespace(**(values | changes))


def test_supplied_workbook_example_and_sources():
    result = calculate(SOURCE, request(), carrier="HMM")
    assert result["totals"] == {"USD": 890.0}
    assert result["rateId"] == 1
    assert result["rateValidUntil"] == "2026-10-15"
    assert result["items"][0]["sourceRow"] == 5


def test_date_selects_later_rate_not_today():
    result = calculate(SOURCE, request(cargoReadyDate="2026-10-20"), carrier="HMM")
    assert result["rateId"] == 5
    assert result["totals"] == {"USD": 960.0}


def test_carrier_unspecified_selects_lowest_not_sum():
    result = calculate(SOURCE, request())
    assert result["carrier"] == "SITC"
    assert result["totals"] == {"USD": 870.0}


def test_no_fabricated_lcl_rate():
    with pytest.raises(ValueError, match="요율이 없습니다"):
        calculate(SOURCE, request(pol="KRPUS", pod="SGSIN", containerType="LCL"))


def test_insurance_minimum_and_container_vs_bl_quantities():
    result = calculate(SOURCE, request(invoiceValue=3000), carrier="HMM", count=2)
    amounts = {r["charge_code"]: r["amount"] for r in result["items"]}
    assert amounts["INS"] == 30
    assert amounts["DOC"] == 30
    assert amounts["OCEAN_FREIGHT"] == 900
    assert amounts["OTHC"] == 240


def test_insurance_percent_above_minimum():
    result = calculate(SOURCE, request(invoiceValue=100000), carrier="HMM")
    assert next(r["amount"] for r in result["items"] if r["charge_code"] == "INS") == 200
