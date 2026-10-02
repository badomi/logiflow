"""제공된 샘플 요율표를 조회하여 양식 기반 데모 XLSX 생성 (실제 케이스/DB 변경 없음).

python -m app.sample_rate_quote
요율표의 조회예시와 동일한 HMM/CNSHA→KRPUS/20GP/CIF/2026-10-05 조건.
공통 입력 필드는 models.py 이름 유지. carrier/count는 조회용 보조값.
"""

import io
import json
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from .config import KST, SERVER_DIR
from .quotation import render_xlsx

SOURCE = SERVER_DIR.parent / "docs" / "양식" / "샘플_요율표.xlsx"
OUTPUT = SERVER_DIR.parent / "output" / "sample-quote"
CONTAINERS = {"20GP": "20FT GP", "40GP": "40FT GP", "40HQ": "40HQ"}


def _rows(wb, sheet):
    ws = wb[sheet]
    headers = [c.value for c in ws[4]]
    required = ({"rate_id", "pol", "pod", "container_type", "carrier", "currency", "amount", "valid_from",
                 "valid_to", "etd", "eta", "transit_days", "free_time_dem", "free_time_det"}
                if sheet == "OceanFreight" else
                {"charge_code", "charge_name", "basis", "applies_pol", "applies_pod", "container_type",
                 "incoterms_condition", "currency", "amount", "rate_pct", "min_amount", "valid_from", "valid_to"})
    if required - set(headers):
        raise ValueError(f"{sheet} 필수 컬럼 누락: {', '.join(sorted(required - set(headers)))}")
    if any(headers.count(key) != 1 for key in required):
        raise ValueError(f"{sheet} 컬럼 중복")
    return [{**dict(zip(headers, values)), "sourceRow": row}
            for row, values in enumerate(ws.iter_rows(min_row=5, values_only=True), 5)
            if values[0] is not None]


def _date(value):
    if isinstance(value, datetime):
        return value.date()
    return value if isinstance(value, date) else date.fromisoformat(str(value))


def _number(value, label, *, integer=False):
    try:
        number = Decimal(str(value))
        if not number.is_finite() or number < 0 or (integer and number != number.to_integral_value()):
            raise ValueError()
        return number
    except (ArithmeticError, ValueError):
        raise ValueError(f"요율표 {label}: 0 이상의 {'정수' if integer else '숫자'}가 필요합니다.") from None


def _schedule(row):
    days = int(_number(row["transit_days"], "transit_days", integer=True))
    etd = _date(row["etd"])
    eta = etd + timedelta(days=days)
    if row["eta"] is not None and _date(row["eta"]) != eta:
        raise ValueError(f"OceanFreight 행 {row['sourceRow']}: ETA와 ETD+운송일수가 다릅니다.")
    return {"etd": etd.isoformat(), "eta": eta.isoformat(), "transitTime": days,
            "freeTimeDem": int(_number(row["free_time_dem"], "free_time_dem", integer=True)),
            "freeTimeDet": int(_number(row["free_time_det"], "free_time_det", integer=True))}


def calculate(path, qi, *, carrier=None, count=1):
    """日付/구간/장비 매칭. 미지정 선사는 최저가, 비율 비용은 최저요금 반영.

    이 어댑터는 샘플 파일에 명시된 CIF 조회예시를 지원한다.
    다른 거래조건의 비용 부담 규칙을 추정하지 않는다.
    """
    if qi.incoterms != "CIF":
        raise ValueError("이 샘플 조회 경로는 CIF 예시만 지원합니다.")
    if count < 1 or int(count) != count:
        raise ValueError("컨테이너 수량은 양의 정수여야 합니다.")
    on_date = _date(qi.cargoReadyDate)
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        ocean, charges = _rows(wb, "OceanFreight"), _rows(wb, "Surcharges")
    finally:
        wb.close()

    def valid(row):
        return _date(row["valid_from"]) <= on_date <= _date(row["valid_to"])

    def matches(value, actual):
        return value in (None, "ALL") or value == actual

    candidates = [r for r in ocean if r["pol"] == qi.pol and r["pod"] == qi.pod
                  and CONTAINERS.get(r["container_type"]) == qi.containerType and valid(r)
                  and (carrier is None or r["carrier"] in (carrier, "ANY"))]
    if not candidates:
        raise ValueError("제공 요율표에 해당 구간·컨테이너·날짜·선사 요율이 없습니다.")
    for row in candidates:
        _schedule(row)
        _number(row["amount"], "amount")
    candidates = [r for r in candidates if on_date <= _date(r["etd"]) <= _date(r["valid_to"])]
    if not candidates:
        raise ValueError("화물 준비일 이후 출항 가능한 유효 스케줄이 없습니다. 요율표의 ETD를 확인하세요.")
    if len({r["currency"] for r in candidates}) > 1:
        raise ValueError("통화가 다른 선사 요율은 환율 없이 최저가 비교할 수 없습니다.")
    chosen = min(candidates, key=lambda r: (Decimal(str(r["amount"])), r["rate_id"]))
    items = []
    valid_to = [_date(chosen["valid_to"])]

    def append(code, label, basis, rate, qty, currency, sheet, row, note):
        amount = (_number(rate, f"{code} 단가") * _number(qty, f"{code} 수량")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        items.append(dict(charge_code=code, charge_label=label, basis=basis, rate=float(rate),
                          qty=qty, amount=float(amount), currency=currency, remark=note,
                          sourceSheet=sheet, sourceRow=row))

    append("OCEAN_FREIGHT", "OCEAN FREIGHT", "PER_CNTR", chosen["amount"], count,
           chosen["currency"], "OceanFreight", chosen["sourceRow"], f"{chosen['carrier']} / PER CNTR")
    for r in charges:
        if not (valid(r) and matches(r["applies_pol"], qi.pol) and matches(r["applies_pod"], qi.pod)
                and (r["container_type"] in (None, "ALL") or CONTAINERS.get(r["container_type"]) == qi.containerType)
                and matches(r["incoterms_condition"], qi.incoterms)):
            continue
        basis = r["basis"]
        if basis == "PERCENT_OF_VALUE":
            if qi.ccy != r["currency"]:
                raise ValueError("보험 계산 통화와 화물가액 통화가 달라 환율이 필요합니다.")
            rate = max(_number(qi.invoiceValue, "invoiceValue") * _number(r["rate_pct"], "rate_pct"),
                       _number(r["min_amount"] or 0, "min_amount"))
            qty, mapped = 1, "PER_BL"
            note = f"VALUE x {float(r['rate_pct']):.1%}; MIN {r['currency']} {r['min_amount']}"
        elif basis in ("PER_CONTAINER", "PER_BL", "PER_SHIPMENT"):
            rate = r["amount"]
            qty = count if basis == "PER_CONTAINER" else 1
            mapped = "PER_CNTR" if basis == "PER_CONTAINER" else "PER_BL"
            note = basis
        else:
            raise ValueError(f"미지원 부과 기준: {basis}")
        if any(item["charge_code"] == r["charge_code"] for item in items):
            raise ValueError(f"적용 부대비용 중복: {r['charge_code']}. 요율표를 확인하세요.")
        append(r["charge_code"], r["charge_name"].upper(), mapped, rate, qty, r["currency"],
               "Surcharges", r["sourceRow"], note)
        valid_to.append(_date(r["valid_to"]))
    missing = {"OTHC", "DTHC", "INS"} - {item["charge_code"] for item in items}
    if missing:
        raise ValueError(f"CIF 샘플 필수 비용 누락: {', '.join(sorted(missing))}")
    totals = {}
    for item in items:
        currency = item["currency"]
        totals[currency] = totals.get(currency, Decimal(0)) + Decimal(str(item["amount"]))
    return {"items": items, "totals": {k: float(v) for k, v in totals.items()},
            "rateValidUntil": min(valid_to).isoformat(), "carrier": chosen["carrier"], "rateId": chosen["rate_id"],
            "schedule": _schedule(chosen)}


def main():
    qi = SimpleNamespace(pol="CNSHA", pod="KRPUS", containerType="20FT GP", incoterms="CIF",
                         cargoReadyDate="2026-10-05", invoiceValue=15000, ccy="USD",
                         customerName="샘플 요율 조회예시", contactName="미지정", qty=None, qtyUnit="",
                         boxL=None, boxW=None, boxH=None, totalCbm=None, grossWeightKg=None,
                         transitTime=None, freeTimeDet=None, freeTimeDem=None)
    result = calculate(SOURCE, qi, carrier="HMM", count=1)
    assert result["totals"] == {"USD": 890.0}, result
    case = SimpleNamespace(quote_input=qi, case_id="SAMPLE-HMM-20GP")
    quote = SimpleNamespace(items=[SimpleNamespace(**r) for r in result["items"]],
                            totals=[SimpleNamespace(currency=k, total=v) for k, v in result["totals"].items()],
                            valid_until=result["rateValidUntil"], created_at=datetime.now(KST), version_no=1)
    wb = load_workbook(io.BytesIO(render_xlsx(case, quote, count=1)))
    ws = wb.active
    ws["B9"] = "수 신 : 샘플 요율 조회예시 / 담당자 미지정"
    ws["B10"] = "발 신 : 레오나해운항공㈜ / 견적 담당"
    ws["B12"] = "Subject : SAMPLE - SHANGHAI / BUSAN (CIF)"
    ws["B14"] = "* TEST QUOTATION ONLY - Not valid for booking or billing."
    ws["B15"] = "* Based on the supplied sample rate workbook; cargo details not provided."
    ws["C17"] = ": SHANGHAI (CNSHA) - BUSAN (KRPUS)"
    ws["C19"] = ": 20FT GP x 1 / HMM; commodity, packages and weight: NOT PROVIDED"
    ws["B20"] = "** INCOTERMS"
    ws["B22"] = "FREIGHT & CHARGES"
    ws["B35"] = "TOTAL (USD)"
    ws["C35"] = "=SUM(E23:E34)"
    ws["C35"].number_format = '"US$"#,##0.00'
    for row in range(23, 23 + len(result["items"])):
        ws[f"E{row}"] = f"=ROUND(C{row}*D{row},2)"
    remarks = [
        "* SOURCE: 샘플_요율표.xlsx / OceanFreight + Surcharges + 조회예시",
        "* CARGO READY DATE: 2026-10-05 / INVOICE VALUE: USD 15,000.00",
        "* ETD / ETA: NOT PROVIDED / TRANSIT TIME: NOT PROVIDED",
        "* FREE TIME (DET / DEM): NOT PROVIDED",
        f"* RATE VALID UNTIL: {result['rateValidUntil']} (selected rates; quote validity not provided)",
        "* PAYMENT TERM: NOT PROVIDED / Currency: USD; no FX conversion applied.",
        "* Charges may vary with FX and cargo details; confirm before shipment.",
        "* SAMPLE-HMM-20GP / Rev. 1 - Separate example; not case LQ-2026-1002-003.",
        "* Insurance: USD 15,000 x 0.2% = USD 30 (minimum USD 30).",
    ]
    for row, remark in enumerate(remarks, 39):
        ws[f"B{row}"] = remark
        ws[f"B{row}"].font = Font(name="맑은 고딕", size=9)
    proof = wb.create_sheet("산출근거")
    proof.append(["샘플_요율표.xlsx 조회예시 대조 — 실제 견적 아님"])
    proof.append(["표준 필드", "값", "요율표 값"])
    for name in ("pol", "pod", "containerType", "incoterms", "cargoReadyDate", "invoiceValue", "ccy"):
        proof.append([name, getattr(qi, name), "20GP" if name == "containerType" else None])
    proof.append(["carrier (조회 보조값)", "HMM"])
    proof.append(["containerCount (조회 보조값)", 1])
    proof.append([])
    proof.append(["charge_code", "charge_name", "source_sheet", "Excel 행", "단가 USD", "수량", "금액 USD", "산정 기준"])
    for item in result["items"]:
        proof.append([item["charge_code"], item["charge_label"], item["sourceSheet"], item["sourceRow"],
                      item["rate"], item["qty"], item["amount"], item["remark"]])
    proof.append(["TOTAL", None, None, None, None, None, 890])
    for cell in proof[13]:
        cell.fill = PatternFill("solid", fgColor="163B66")
        cell.font = Font(color="FFFFFF", bold=True)
    for col, width in {"A":32,"B":30,"C":22,"D":12,"E":16,"F":10,"G":16,"H":42}.items():
        proof.column_dimensions[col].width = width
    proof.freeze_panes = "A14"
    proof.sheet_properties.pageSetUpPr.fitToPage = True
    proof.page_setup.orientation = "landscape"
    proof.page_setup.paperSize = proof.PAPERSIZE_A4
    proof.page_setup.fitToWidth = 1
    proof.page_setup.fitToHeight = 1
    proof.print_area = f"A1:H{proof.max_row}"
    OUTPUT.mkdir(parents=True, exist_ok=True)
    target = OUTPUT / "샘플견적서_상하이-부산_HMM_20GP.xlsx"
    wb.save(target)
    (OUTPUT / "calculation.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(target)
    print(result["totals"])


if __name__ == "__main__":
    main()
