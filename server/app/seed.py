"""마스터 데이터 등록.

    python -m app.seed               컨테이너 마스터만 등록
    python -m app.seed --sample-rates + 발주 측 예시 견적서 금액을 '테스트용 요율'로 등록

⚠ 테스트용 요율 출처 (모두 발주 측 제공 자료, 실제 요율 아님):
  1) docs/양식의 LCL·FOB 예시 견적서 — 인천 LCL 부대비용, 부산 40HQ 부대비용
  2) docs/samples.md 3번 'D선사 월간 운임표' — 부산 출발 해상운임 17개 항구 × 20'GP·40'HC·40'REEFER
     (자료에 출발항이 없어 부산으로 가정, VALIDITY '~월말'은 월 표기가 없어 기간 없이 등록)
  발주 측 요율표·부대비용 Excel(미수령)을 받으면 그 값으로 교체해야 한다.
"""

import argparse

from sqlalchemy import delete
from sqlalchemy.orm import Session

from .db import SessionLocal, init_db
from .models import ContainerType, LaneRule, Rate

CONTAINERS = [
    ("LCL", "LCL 혼재 화물 (R/T 청구)", None, None),
    ("20FT GP", "20ft 일반 컨테이너", 21_700, 33.0),
    ("40FT GP", "40ft 일반 컨테이너", 26_500, 67.0),
    ("40HQ", "40ft 하이큐브 컨테이너", 26_500, 76.0),
    ("40RF", "40ft 냉동·냉장(Reefer) 컨테이너", None, None),  # 허용중량은 자료 없음 → 초과 검사 안 함
]

SAMPLE_SOURCE = "TEST: docs/양식 예시 견적서 (실제 요율 아님)"

# (charge_code, 견적서 표기, pol, pod, container, incoterms, basis, 단가, 통화, 비고, 순서)
SAMPLE_RATES = [
    # LCL견적서_예시.pdf — INCHEON → SYDNEY, LCL, CIF
    ("OCEAN_FREIGHT", "OCEAN FREIGHT", "KRINC", "AUSYD", "LCL", "CIF", "PER_RT", 500.0, "USD", "PER R/T", 10),
    ("THC", "THC", "KRINC", None, "LCL", None, "PER_RT", 5_500, "KRW", "PER R/T", 20),
    ("CFS", "CFS", "KRINC", None, "LCL", None, "PER_RT", 6_500, "KRW", "PER R/T", 30),
    ("WFG", "WFG", "KRINC", None, "LCL", None, "PER_RT", 203, "KRW", "PER R/T", 40),
    ("SHUTTLE", "SHUTTLE CHARGE", "KRINC", None, "LCL", None, "PER_RT", 5_000, "KRW", "PER R/T", 50),
    ("DOC_FEE", "DOCUMENT FEE", "KRINC", None, "LCL", None, "PER_BL", 90_000, "KRW", "PER B/L", 60),
    # 예시 견적서는 통관비·보험료를 최소 금액(MIN)으로 청구 — 실제 산정식은 비고에 표기
    ("CUSTOMS", "CUSTOMS CLEARANCE FEE", "KRINC", None, "LCL", None, "PER_BL", 15_000, "KRW", "INV.V x 1/1,000 (MIN 기준)", 80),
    ("INSURANCE", "INSURANCE FEE", "KRINC", None, "LCL", "CIF", "PER_BL", 18_000, "KRW", "INV.V x 110% x 보험요율 (MIN 기준)", 90),
    # FOB견적서_예시.pdf — BUSAN, 40HQ, FOB (해상운임 없음: FOB는 매수인 부담)
    ("THC", "THC", "KRPUS", None, "40HQ", None, "PER_CNTR", 410_000, "KRW", "PER CNTR", 20),
    ("WFG", "WFG", "KRPUS", None, "40HQ", None, "PER_CNTR", 10_220, "KRW", "PER CNTR", 40),
    ("SEAL", "SEAL CHARGE", "KRPUS", None, "40HQ", None, "PER_CNTR", 8_000, "KRW", "PER CNTR", 45),
    ("DOC_FEE", "DOCUMENT FEE", "KRPUS", None, "40HQ", None, "PER_BL", 70_000, "KRW", "PER B/L", 60),
    ("CUSTOMS", "CUSTOMS CLEARANCE FEE", "KRPUS", None, "40HQ", None, "AT_COST", None, "KRW", "INV.V x 1/1,000", 80),
    ("INSURANCE", "INSURANCE FEE", "KRPUS", None, "40HQ", None, "AT_COST", None, "KRW", "INV.V x 110% x 보험요율", 90),
]
# docs/samples.md 3번 D선사 월간 운임표 (USD / 컨테이너). None = 표의 '-' (취급 안 함 → 등록하지 않음)
OFT_SOURCE = "TEST: samples.md 3번 D선사 월간 운임표 (교육용 가공 금액, POL 부산 가정)"
MONTHLY_OFT = [
    # (UN/LOCODE, 20'GP, 40'HC, 40'REEFER, REMARK)
    ("MYPKG", 500, 750, 1_450, "INCLUSIVE ISPS(DIRECT)"),
    ("SGSIN", 500, 750, 1_450, "INCLUSIVE ISPS(DIRECT)"),
    ("INMAA", 2_650, 2_850, 3_850, "Subj to ISPS(DIRECT)"),
    ("INKAT", 2_650, 2_850, 3_850, "Subj to ISPS(DIRECT)"),
    ("INNSA", 3_400, 3_650, 4_400, "Subj to ISPS(DIRECT)"),
    ("INPAV", 3_400, 3_650, 4_400, "Subj to ISPS(DIRECT)"),
    ("INMUN", 3_400, 3_650, 4_400, "Subj to ISPS(DIRECT)"),
    ("PKKHI", 3_500, 3_750, 4_600, "Subj to ISPS(DIRECT)"),
    ("AEJEA", 7_350, 9_550, None, "Subj to ISPS"),
    ("AEKLF", 6_950, 9_150, None, "Subj to ISPS(DIRECT)"),
    ("OMSOH", 6_950, 9_150, None, "Subj to ISPS(DIRECT)"),
    ("SAJED", 7_700, 9_900, None, "Subj to ISPS(T/S)"),
    ("SAKAC", 7_150, 8_800, None, "Subj to ISPS(T/S)"),
    ("EGSOK", 7_700, 9_900, None, "Subj to ISPS(T/S)"),
    ("JOAQJ", 7_700, 9_900, None, "Subj to ISPS(T/S)"),
    ("SYLTK", 6_600, 8_800, None, "Subj to ISPS(T/S)"),
    ("TRMER", 6_050, 7_700, None, "Subj to ISPS(T/S)"),
]
# 하단 조건 'SURCHARGES : ISPS USD 0 (Common Item)' → 별도 비용 줄 없음

# 내륙운송비(TRUCKING)는 픽업지마다 달라 6장 필드만으로 고를 수 없다 → 요율표 수령 후 픽업지 기준 설계 필요

SAMPLE_LANES = [
    ("KRINC", "AUSYD", "LCL", "USD", None, None, None, 14),
    ("KRPUS", None, "40HQ", "USD", None, None, None, 14),
]


def seed_containers(session: Session) -> None:
    for code, label, max_kg, max_cbm in CONTAINERS:
        session.merge(ContainerType(code=code, label=label, max_gross_kg=max_kg, max_cbm=max_cbm))


def seed_sample_rates(session: Session) -> None:
    session.execute(delete(Rate).where(Rate.source.in_([SAMPLE_SOURCE, OFT_SOURCE])))
    for pod, gp20, hc40, rf40, remark in MONTHLY_OFT:
        for cntr, price in (("20FT GP", gp20), ("40HQ", hc40), ("40RF", rf40)):
            if price is None:
                continue
            # 해상운임은 CIF 견적에만 들어간다 (FOB는 매수인 부담)
            session.add(Rate(charge_code="OCEAN_FREIGHT", charge_label="OCEAN FREIGHT", pol="KRPUS", pod=pod,
                             container_type=cntr, incoterms="CIF", basis="PER_CNTR", unit_price=float(price),
                             currency="USD", remark=f"PER CNTR · {remark}", sort_no=10, source=OFT_SOURCE))
    for code, label, pol, pod, cntr, terms, basis, price, ccy, remark, order in SAMPLE_RATES:
        session.add(Rate(
            charge_code=code, charge_label=label, pol=pol, pod=pod, container_type=cntr, incoterms=terms,
            basis=basis, unit_price=price, currency=ccy, remark=remark, sort_no=order, source=SAMPLE_SOURCE,
        ))
    for pol, pod, cntr, ccy, transit, det, dem, validity in SAMPLE_LANES:
        session.execute(delete(LaneRule).where(
            LaneRule.pol == pol, LaneRule.container_type == cntr,
            LaneRule.pod.is_(None) if pod is None else LaneRule.pod == pod,
        ))
        session.add(LaneRule(pol=pol, pod=pod, container_type=cntr, quote_currency=ccy, transit_days=transit,
                             free_time_det=det, free_time_dem=dem, validity_days=validity))


def main() -> None:
    parser = argparse.ArgumentParser(description="마스터 데이터 등록")
    parser.add_argument("--sample-rates", action="store_true", help="예시 견적서 금액을 테스트용 요율로 등록")
    args = parser.parse_args()
    init_db()
    with SessionLocal() as session:
        seed_containers(session)
        if args.sample_rates:
            seed_sample_rates(session)
        session.commit()
    print("등록 완료" + (" (테스트용 요율 포함 — 실제 요율 아님)" if args.sample_rates else ""))


if __name__ == "__main__":
    main()
