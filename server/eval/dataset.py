"""정밀도 측정(NFR-01) 평가 세트 — 메일과 필드별 정답.

- 'eml'은 발주 측 샘플(server/samples/eml), 나머지는 B트랙 입력 특성을 담아 직접 만든 변형 메일이다.
- 정답(expected)에 없는 필드는 '메일에 없음(None)'이 정답이다 → 시스템이 채우면 오추출로 센다.
- 값 기준: 6장 형식(단위 환산 후). 메일에 적혔지만 애매한 값(단위 없는 규격 등)도 실제 값을 정답으로 둔다
  — 시스템이 안 채우면 정밀도에는 영향 없고 채움률만 떨어진다(정의서: 미추출은 오류가 아니라 보완 요청).
- 발주 측 샘플을 바탕으로 한 자료이므로 외부 재배포 금지.
"""

import json
from pathlib import Path

CASES_DIR = Path(__file__).resolve().parent / "cases"

LCL_01 = {
    "commodity": "스프링노트(일반 문구류)", "qty": 10, "qtyUnit": "CTNS", "packing": "Carton",
    "boxL": 310, "boxW": 450, "boxH": 270, "totalCbm": 0.19, "pod": "AUSYD", "containerType": "LCL",
    "grossWeightKg": 500, "contactEmail": "shipper@example.com",
}

CASES = [
    # ---------------------------------------------------------------- 발주 측 샘플
    {"id": "S01", "desc": "8-1 LCL 요청 (선적항 미정, 조건 2개, 픽업지≠선적항)",
     "mails": [{"eml": "01_lcl_request.eml"}], "expected": LCL_01},
    {"id": "S01+03", "desc": "8-1 LCL + 보완 회신 병합 (인용문 포함)",
     "mails": [{"eml": "01_lcl_request.eml"}, {"eml": "03_lcl_reply.eml", "role": "REPLY"}],
     "expected": {**LCL_01, "cargoReadyDate": "2026-10-20", "invoiceValue": 3000, "ccy": "USD"}},
    {"id": "S06", "desc": "8-1 LCL (EUC-KR 인코딩)", "mails": [{"eml": "06_lcl_request_euckr.eml"}], "expected": LCL_01},
    {"id": "S02", "desc": "8-2 FOB 다품목 → 수동 처리", "mails": [{"eml": "02_fob_multi_item.eml"}],
     "expected": {"contactEmail": "shipper@example.com"}, "multipleItems": True},

    # ---------------------------------------------------------------- 변형 메일
    {"id": "V01", "desc": "국문 라벨 FCL FOB + 서명(㈜·직함)",
     "mails": [{"from": ("김민수", "ms.kim@hanbit-trade.co.kr"), "subject": "[견적요청] 부산 → 상하이 40HQ FOB", "body": """안녕하세요, 한빛무역 김민수입니다.
아래 건 해상 운임 견적 부탁드립니다.

품목: 플라스틱 부품
수량: 500박스
박스 규격: 600 × 400 × 400mm
총 부피: 48CBM
총중량: 10,000kg
선적항: 부산
도착항: 상하이
화물 준비일: 2026-11-05
조건: FOB
컨테이너: 40HQ 1대
인보이스 금액: USD 15,000
결제조건: T/T 30% Advance

감사합니다.
김민수 대리
㈜한빛무역 | 해외영업팀
Tel. 051-123-4567"""}],
     "expected": {"commodity": "플라스틱 부품", "qty": 500, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 600, "boxW": 400,
                  "boxH": 400, "totalCbm": 48, "pol": "KRPUS", "pod": "CNSHA", "cargoReadyDate": "2026-11-05",
                  "incoterms": "FOB", "containerType": "40HQ", "grossWeightKg": 10000, "invoiceValue": 15000, "ccy": "USD",
                  "paymentTerm": "T/T 30% Advance", "customerName": "㈜한빛무역", "contactName": "김민수 대리",
                  "contactEmail": "ms.kim@hanbit-trade.co.kr"}},

    {"id": "V02", "desc": "영문 LCL, LBS·CFT·inch 환산 (FR-203)",
     "mails": [{"from": ("John Smith", "john.smith@abc-trading.com"), "subject": "RFQ - LCL Busan to Los Angeles", "body": """Dear LEONA team,

Please quote LCL ocean freight for the shipment below.

Commodity: Kitchen Utensils
Quantity: 40 cartons
Carton Size: 20 x 16 x 12 inch
Total Volume: 88.9 CFT
Gross Weight: 1,320 LBS
POL: Busan, Korea
POD: Los Angeles, USA
Cargo Ready Date: Nov 12, 2026
Incoterms: CIF
Invoice Value: USD 8,400
Payment Terms: L/C at sight

Best regards,
John Smith
Purchasing Manager
ABC Trading Co., Ltd."""}],
     "expected": {"commodity": "Kitchen Utensils", "qty": 40, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 508, "boxW": 406,
                  "boxH": 305, "totalCbm": 2.5174, "pol": "KRPUS", "pod": "USLAX", "cargoReadyDate": "2026-11-12",
                  "incoterms": "CIF", "containerType": "LCL", "grossWeightKg": 598.742, "invoiceValue": 8400, "ccy": "USD",
                  "paymentTerm": "L/C at sight", "customerName": "ABC Trading Co., Ltd.", "contactName": "John Smith",
                  "contactEmail": "john.smith@abc-trading.com"}},

    {"id": "V03", "desc": "국·영문 혼용 라벨, cm 규격, 단위 없는 CBM (FR-205)",
     "mails": [{"from": ("이지은", "jieun.lee@seoulcos.kr"), "subject": "인천-시드니 LCL 견적 문의", "body": """안녕하세요.
LCL 견적 요청드립니다.

- Commodity : 화장품 (기초 스킨케어)
- Q'ty : 25 CTNS
- Dimension : 50 x 40 x 30 cm
- CBM : 1.5
- G.W : 380 KGS
- POL : 인천
- POD : SYDNEY
- Cargo Ready : 2026.10.28
- Term : CIF
- Invoice : USD 12,500
- Payment : T/T in advance

감사합니다.
이지은 과장
서울코스메틱 주식회사"""}],
     "expected": {"commodity": "화장품 (기초 스킨케어)", "qty": 25, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 500,
                  "boxW": 400, "boxH": 300, "totalCbm": 1.5, "pol": "KRINC", "pod": "AUSYD", "cargoReadyDate": "2026-10-28",
                  "incoterms": "CIF", "containerType": "LCL", "grossWeightKg": 380, "invoiceValue": 12500, "ccy": "USD",
                  "paymentTerm": "T/T in advance", "customerName": "서울코스메틱 주식회사", "contactName": "이지은 과장",
                  "contactEmail": "jieun.lee@seoulcos.kr"}},

    {"id": "V04", "desc": "Outlook 번호 목록 갈라짐 (1.\\r\\n내용)",
     "mails": [{"from": ("박준호", "junho.park@daehan-steel.com"), "subject": "부산→호치민 20FT FOB 견적",
                "body": "안녕하세요. 아래 조건으로 견적 부탁드립니다.\r\n  1.\r\n품목: 철강 볼트\r\n  2.\r\n수량: 10 팔레트\r\n  3.\r\n"
                        "총중량: 8,000 kg\r\n  4.\r\n전체 부피: 12 CBM\r\n  5.\r\n선적항: 부산항\r\n  6.\r\n도착항: 호치민\r\n  7.\r\n"
                        "출고 가능일: 2026년 11월 20일\r\n  8.\r\n운송조건: FOB\r\n  9.\r\n컨테이너: 20피트 1대\r\n감사합니다.\r\n"
                        "박준호 차장\r\n대한스틸㈜"}],
     "expected": {"commodity": "철강 볼트", "qty": 10, "qtyUnit": "PLTS", "packing": "Pallet", "totalCbm": 12,
                  "pol": "KRPUS", "pod": "VNSGN", "cargoReadyDate": "2026-11-20", "incoterms": "FOB",
                  "containerType": "20FT GP", "grossWeightKg": 8000, "customerName": "대한스틸㈜", "contactName": "박준호 차장",
                  "contactEmail": "junho.park@daehan-steel.com"}},

    {"id": "V05", "desc": "라벨 없는 문장형 메일 (룰 약점 → LLM 대상), 연도 없는 날짜",
     "mails": [{"from": ("최서연", "seoyeon@greenfood.kr"), "subject": "견적 문의드립니다", "body": """안녕하세요, 그린푸드 최서연입니다.
다음 달 초에 부산에서 싱가포르로 20피트 컨테이너 한 대 보낼 예정이라 FOB 기준 견적 부탁드립니다.
물건은 김 스낵이고 800박스, 전체 무게는 6톤 정도 됩니다. 박스 하나 크기는 40 x 30 x 25cm예요.
인보이스 금액은 대략 2만 달러입니다.
감사합니다.
최서연 드림"""}],
     "expected": {"commodity": "김 스낵", "qty": 800, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 400, "boxW": 300,
                  "boxH": 250, "pol": "KRPUS", "pod": "SGSIN", "incoterms": "FOB", "containerType": "20FT GP",
                  "grossWeightKg": 6000, "invoiceValue": 20000, "ccy": "USD", "contactName": "최서연",
                  "contactEmail": "seoyeon@greenfood.kr"}},

    {"id": "V06", "desc": "영문 DDP + Gmail 영문 인용 회신, 회신에서 중량 정정(문장형)",
     "mails": [
         {"from": ("Mike Chen", "mike.chen@pacificsupply.com"), "subject": "Quote request: Busan - Rotterdam 40GP DDP", "body": """Hello,

Please provide a DDP quotation.

Commodity: Auto Parts
Packages: 20 pallets
Gross Weight: 12.5 MT
Volume: 28 m3
POL: Busan
POD: Rotterdam
Container: 40GP x 1

Regards,
Mike Chen
Logistics Coordinator
Pacific Supply Inc."""},
         {"role": "REPLY", "from": ("Mike Chen", "mike.chen@pacificsupply.com"),
          "subject": "Re: [LQ-2026-0929-005] Quote request: Busan - Rotterdam 40GP DDP", "body": """Hi,

Please see below.
Cargo ready date: 2026-12-01
Invoice value: EUR 45,000
Payment terms: T/T 30 days
Gross weight is revised to 13,000 KGS.

Regards,
Mike Chen

On Tue, Sep 29, 2026 at 11:46 AM LEONA <quote@leona.example.com> wrote:
> Gross Weight: 12.5 MT
> Please advise cargo ready date."""}],
     "expected": {"commodity": "Auto Parts", "qty": 20, "qtyUnit": "PLTS", "packing": "Pallet", "totalCbm": 28,
                  "pol": "KRPUS", "pod": "NLRTM", "cargoReadyDate": "2026-12-01", "incoterms": "DDP",
                  "containerType": "40FT GP", "grossWeightKg": 13000, "invoiceValue": 45000, "ccy": "EUR",
                  "paymentTerm": "T/T 30 days", "customerName": "Pacific Supply Inc.", "contactName": "Mike Chen",
                  "contactEmail": "mike.chen@pacificsupply.com"}},

    {"id": "V07", "desc": "정보가 거의 없는 문의 (지어내지 않는지)",
     "mails": [{"from": ("정우성", "ws.jung@example.co.kr"), "subject": "운임 문의", "body": """안녕하세요.
의류 제품 부산에서 LA로 보내려고 하는데 대략적인 운임 알 수 있을까요?
수량이나 무게는 아직 정해지지 않았습니다.
감사합니다."""}],
     "expected": {"commodity": "의류 제품", "pol": "KRPUS", "pod": "USLAX", "contactName": "정우성",
                  "contactEmail": "ws.jung@example.co.kr"}},

    {"id": "V08", "desc": "영문 다품목 (Item 1 / Item 2) → 수동 처리 (FR-210)",
     "mails": [{"from": ("Anna Lee", "anna@brightlight.com"), "subject": "RFQ Busan-Hamburg FCL", "body": """Hello, please quote for two items in one 40HQ.

Item 1: LED Lamps
Quantity: 200 cartons
Carton size: 50 x 40 x 30 cm
Item 2: Lamp Stands
Quantity: 80 cartons
Carton size: 120 x 30 x 30 cm
POL: Busan
POD: Hamburg
Incoterms: FOB

Thanks,
Anna Lee
Bright Light Co., Ltd."""}],
     "expected": {"customerName": "Bright Light Co., Ltd.", "contactName": "Anna Lee", "contactEmail": "anna@brightlight.com"},
     "multipleItems": True},

    {"id": "V09", "desc": "수입 EXW/LCL, 발주 측 양식식 'CARGO DETAIL' 한 줄",
     "mails": [{"from": ("한서진", "sj.han@fmanufacturing.co.kr"), "subject": "해상 수입 견적 요청 (EXW / LCL)", "body": """안녕하세요.
미국 창고 픽업 건 해상 수입 견적 부탁드립니다.
조건: EXW / LCL
POL: LONG BEACH, USA
POD: BUSAN, KOREA
CARGO DETAIL: 3PKGS / 1,200KGS / 3.0CBM (stackable)
화물 준비일: 2026-10-15
물품 가액: USD 25,000
감사합니다.
한서진 대리
F제조 주식회사"""}],
     "expected": {"qty": 3, "qtyUnit": "PKGS", "totalCbm": 3.0, "pol": "USLGB", "pod": "KRPUS",
                  "cargoReadyDate": "2026-10-15", "incoterms": "EXW", "containerType": "LCL", "grossWeightKg": 1200,
                  "invoiceValue": 25000, "ccy": "USD", "customerName": "F제조 주식회사", "contactName": "한서진 대리",
                  "contactEmail": "sj.han@fmanufacturing.co.kr"}},

    {"id": "V10", "desc": "전각 콜론(：)·곱하기(×)·20' GP·달러 표기",
     "mails": [{"from": ("오하늘", "haneul.oh@moonlight.kr"), "subject": "광양 → 요코하마 견적 요청 (20ft)", "body": """품목：세라믹 타일
박스 수량：300 CTNS
박스 규격：40 × 40 × 20 cm
총 부피：9.6 CBM
총중량：9,000 KG
선적항：광양
도착항：요코하마
화물 준비일：2026/11/03
거래조건：CIF
컨테이너：20' GP x 1
인보이스 금액：9,800달러
결제조건：D/P
감사합니다.
오하늘 주임
문라이트산업㈜"""}],
     "expected": {"commodity": "세라믹 타일", "qty": 300, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 400, "boxW": 400,
                  "boxH": 200, "totalCbm": 9.6, "pol": "KRKAN", "pod": "JPYOK", "cargoReadyDate": "2026-11-03",
                  "incoterms": "CIF", "containerType": "20FT GP", "grossWeightKg": 9000, "invoiceValue": 9800, "ccy": "USD",
                  "paymentTerm": "D/P", "customerName": "문라이트산업㈜", "contactName": "오하늘 주임",
                  "contactEmail": "haneul.oh@moonlight.kr"}},

    {"id": "V11", "desc": "영문, 단위 없는 크레이트 규격(검토 필요), 톤 단위",
     "mails": [{"from": ("Kevin Park", "kevin@k-furniture.com"), "subject": "Quotation request - Incheon to Melbourne", "body": """Hi,

We need a quote for furniture parts to Melbourne.

Commodity: Furniture Parts
Total packages: 15 crates
Crate size: 1200 x 800 x 900
Total CBM: 12.96
Gross weight: 3.2 tons
Port of loading: Incheon
Destination port: Melbourne
Ready date: 2026-11-30
Incoterm: CIF

Thanks
Kevin Park
Export Manager
K Furniture Inc."""}],
     "expected": {"commodity": "Furniture Parts", "qty": 15, "qtyUnit": "CRATES", "packing": "Crate", "boxL": 1200,
                  "boxW": 800, "boxH": 900, "totalCbm": 12.96, "pol": "KRINC", "pod": "AUMEL", "cargoReadyDate": "2026-11-30",
                  "incoterms": "CIF", "grossWeightKg": 3200, "customerName": "K Furniture Inc.", "contactName": "Kevin Park",
                  "contactEmail": "kevin@k-furniture.com"}},

    {"id": "V12", "desc": "본문 일부 + 첨부 패킹리스트 XLSX (FR-204)",
     "mails": [{"from": ("윤도현", "dh.yoon@hanaelec.co.kr"), "subject": "부산-칭다오 LCL 견적 (패킹리스트 첨부)", "body": """안녕하세요.
첨부 패킹리스트 기준으로 LCL 견적 부탁드립니다.
품목: 전자부품
선적항: 부산
도착항: 칭다오
조건: FOB
화물 준비일: 2026-10-25
감사합니다.
윤도현 과장
하나전자㈜""",
                "attachments": [("packing_list.xlsx", [["Package", "12 CTNS"], ["Carton Size", "60 x 40 x 35 cm"],
                                                       ["Total CBM", "1.008 CBM"], ["Gross Weight", "240 KGS"],
                                                       ["Invoice Value", "USD 6,000"]])]}],
     "expected": {"commodity": "전자부품", "qty": 12, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 600, "boxW": 400,
                  "boxH": 350, "totalCbm": 1.008, "pol": "KRPUS", "pod": "CNTAO", "cargoReadyDate": "2026-10-25",
                  "incoterms": "FOB", "containerType": "LCL", "grossWeightKg": 240, "invoiceValue": 6000, "ccy": "USD",
                  "customerName": "하나전자㈜", "contactName": "윤도현 과장", "contactEmail": "dh.yoon@hanaelec.co.kr"}},
]


# ---------------------------------------------------------------- 부산 기준 기능 테스트 메일 (담당자 테스트용과 같은 내용)
# 이 메일들도 룰을 만든 사람이 썼으므로 '블라인드'는 아니다. MUST-SHIP ① 시연용 샘플 수(20건)를 채우고 LLM 효과를 보는 용도.
_T_LCL = {"commodity": "스프링노트", "qty": 10, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 310, "boxW": 450,
          "boxH": 270, "totalCbm": 0.38, "pol": "KRPUS", "pod": "SGSIN", "cargoReadyDate": "2026-10-20",
          "incoterms": "CIF", "containerType": "LCL", "grossWeightKg": 500, "invoiceValue": 3000, "ccy": "USD",
          "paymentTerm": "T/T 30% Advance", "customerName": "㈜한빛무역", "contactName": "김민수 대리",
          "contactEmail": "ms.kim@hanbit-trade.co.kr"}

CASES += [
    {"id": "T01", "desc": "부산→싱가포르 LCL, 정보 모두 있음",
     "mails": [{"from": ("김민수", "ms.kim@hanbit-trade.co.kr"), "subject": "부산-싱가포르 LCL 해상 견적 요청", "body": """안녕하세요. 아래 조건으로 LCL 견적 부탁드립니다.

품목: 스프링노트
박스 수량: 10박스
박스 1개 규격: 310 × 450 × 270mm
전체 부피: 0.38CBM
예상 총중량: 500kg
선적항: 부산
도착항: 싱가포르
화물 준비일: 2026-10-20
조건: CIF
선적방식: LCL
인보이스 금액: USD 3,000
결제조건: T/T 30% Advance

감사합니다.
김민수 대리
㈜한빛무역"""}], "expected": _T_LCL},

    {"id": "T02", "desc": "정보 부족 요청 + 보완 질문 번호로만 답한 회신",
     "asked": ["pol", "cargoReadyDate", "incoterms", "invoiceValue", "paymentTerm"],
     "mails": [
         {"from": ("이지은", "jieun.lee@seoulmungu.co.kr"), "subject": "싱가포르행 LCL 견적 문의", "body": """안녕하세요, 서울문구 이지은입니다.
싱가포르로 보낼 화물 LCL 운임 견적 부탁드립니다.

품목: 문구류(노트)
박스 수량: 20박스
박스 규격: 40 x 30 x 25 cm
총중량: 180kg
도착항: 싱가포르
선적방식: LCL

감사합니다.
이지은 과장
서울문구 주식회사"""},
         {"role": "REPLY", "from": ("이지은", "jieun.lee@seoulmungu.co.kr"), "subject": "RE: [LQ-2026-0930-002] 견적 보완 요청",
          "body": "1. 부산\n2. 2026-10-27\n3. CIF\n4. USD 2,400\n5. T/T 100% Advance"}],
     "expected": {"commodity": "문구류(노트)", "qty": 20, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 400, "boxW": 300,
                  "boxH": 250, "pol": "KRPUS", "pod": "SGSIN", "cargoReadyDate": "2026-10-27", "incoterms": "CIF",
                  "containerType": "LCL", "grossWeightKg": 180, "invoiceValue": 2400, "ccy": "USD",
                  "paymentTerm": "T/T 100% Advance", "customerName": "서울문구 주식회사", "contactName": "이지은 과장",
                  "contactEmail": "jieun.lee@seoulmungu.co.kr"}},

    {"id": "T03", "desc": "부산→싱가포르 40HQ 2대 CIF",
     "mails": [{"from": ("", "buyer@plasticparts.co.kr"), "subject": "[견적요청] 부산 → 싱가포르 40HQ CIF", "body": """품목: 플라스틱 부품
수량: 1,000박스
총 부피: 96CBM
총중량: 20,000kg
선적항: 부산
도착항: 싱가포르
화물 준비일: 2026-11-05
조건: CIF
컨테이너: 40HQ 2대
인보이스 금액: USD 30,000
결제조건: T/T 30% Advance

감사합니다."""}],
     "expected": {"commodity": "플라스틱 부품", "qty": 1000, "qtyUnit": "CTNS", "packing": "Carton", "totalCbm": 96,
                  "pol": "KRPUS", "pod": "SGSIN", "cargoReadyDate": "2026-11-05", "incoterms": "CIF",
                  "containerType": "40HQ", "grossWeightKg": 20000, "invoiceValue": 30000, "ccy": "USD",
                  "paymentTerm": "T/T 30% Advance", "contactEmail": "buyer@plasticparts.co.kr"}},

    {"id": "T04", "desc": "부산→포트클랑 40HQ, 대수 없음",
     "mails": [{"from": ("", "import@furniture.co.kr"), "subject": "부산 → 포트클랑 FCL 견적", "body": """품목: 가구 부품
수량: 800박스
총 부피: 100CBM
총중량: 20,000kg
선적항: 부산
도착항: 포트클랑
화물 준비일: 2026-11-10
조건: CIF
컨테이너: 40HQ
인보이스 금액: USD 40,000
결제조건: L/C at sight

감사합니다."""}],
     "expected": {"commodity": "가구 부품", "qty": 800, "qtyUnit": "CTNS", "packing": "Carton", "totalCbm": 100,
                  "pol": "KRPUS", "pod": "MYPKG", "cargoReadyDate": "2026-11-10", "incoterms": "CIF",
                  "containerType": "40HQ", "grossWeightKg": 20000, "invoiceValue": 40000, "ccy": "USD",
                  "paymentTerm": "L/C at sight", "contactEmail": "import@furniture.co.kr"}},

    {"id": "T05", "desc": "부산→상하이 40HQ FOB",
     "mails": [{"from": ("박준호", "junho.park@daehan-apparel.co.kr"), "subject": "부산 → 상하이 40HQ FOB 견적", "body": """품목: 의류
수량: 1,000박스
박스 규격: 460 × 290 × 250mm
총중량: 4,950kg
선적항: 부산
도착항: 상하이
화물 준비일: 2026-11-10
조건: FOB
컨테이너: 40HQ 1대
인보이스 금액: USD 60,000
결제조건: L/C at sight

감사합니다.
박준호 차장
대한어패럴㈜"""}],
     "expected": {"commodity": "의류", "qty": 1000, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 460, "boxW": 290,
                  "boxH": 250, "pol": "KRPUS", "pod": "CNSHA", "cargoReadyDate": "2026-11-10", "incoterms": "FOB",
                  "containerType": "40HQ", "grossWeightKg": 4950, "invoiceValue": 60000, "ccy": "USD",
                  "paymentTerm": "L/C at sight", "customerName": "대한어패럴㈜", "contactName": "박준호 차장",
                  "contactEmail": "junho.park@daehan-apparel.co.kr"}},

    {"id": "T06", "desc": "라벨 없는 문장형 (말로 쓴 수, 나뉜 금액·통화, 음차 결제조건) — LLM 효과 확인용",
     "mails": [{"from": ("최서연", "seoyeon@greenoffice.kr"), "subject": "견적 부탁드립니다", "body": """안녕하세요, 그린오피스 최서연입니다.
부산에서 싱가포르로 LCL로 보낼 노트가 열 박스 있어요.
한 박스가 가로 310 세로 450 높이 270mm이고 무게는 다 합쳐 오백 킬로 정도예요.
화물은 10월 20일쯤 준비되고 CIF로 부탁드려요.
금액은 3,000이고 통화는 USD입니다.
결제는 티티 30% 선결제로 할게요.

감사합니다.
최서연 드림"""}],
     # 연도 없는 '10월 20일'은 측정일 기준 다가오는 날로 추정한다 (2026-10-27 이후에 재면 2027년이 되어 틀림으로 셈)
     "expected": {"commodity": "노트", "qty": 10, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 310, "boxW": 450,
                  "boxH": 270, "pol": "KRPUS", "pod": "SGSIN", "cargoReadyDate": "2026-10-20", "incoterms": "CIF",
                  "containerType": "LCL", "grossWeightKg": 500, "invoiceValue": 3000, "ccy": "USD",
                  "paymentTerm": "T/T 30% Advance", "customerName": "그린오피스", "contactName": "최서연",
                  "contactEmail": "seoyeon@greenoffice.kr"}},

    {"id": "T07", "desc": "영문 LCL 부산→싱가포르, LBS·CFT·inch",
     "mails": [{"from": ("John Smith", "john.smith@abc-trading.com"), "subject": "RFQ - LCL Busan to Singapore", "body": """Dear LEONA team,

Please quote LCL ocean freight for the shipment below.

Commodity: Kitchen Utensils
Quantity: 40 cartons
Carton Size: 20 x 16 x 12 inch
Total Volume: 88.9 CFT
Gross Weight: 1,320 LBS
POL: Busan, Korea
POD: Singapore
Cargo Ready Date: Nov 12, 2026
Incoterms: CIF
Invoice Value: USD 8,400
Payment Terms: L/C at sight

Best regards,
John Smith
Purchasing Manager
ABC Trading Co., Ltd."""}],
     "expected": {"commodity": "Kitchen Utensils", "qty": 40, "qtyUnit": "CTNS", "packing": "Carton", "boxL": 508, "boxW": 406,
                  "boxH": 305, "totalCbm": 2.5174, "pol": "KRPUS", "pod": "SGSIN", "cargoReadyDate": "2026-11-12",
                  "incoterms": "CIF", "containerType": "LCL", "grossWeightKg": 598.742, "invoiceValue": 8400, "ccy": "USD",
                  "paymentTerm": "L/C at sight", "customerName": "ABC Trading Co., Ltd.", "contactName": "John Smith",
                  "contactEmail": "john.smith@abc-trading.com"}},

    {"id": "T08", "desc": "부산→첸나이 40HQ (월간 운임표 구간)",
     "mails": [{"from": ("", "sales@elec-parts.co.kr"), "subject": "부산 → 첸나이 40HQ 견적", "body": """품목: 전자부품
수량: 300박스
총 부피: 55CBM
총중량: 9,000kg
선적항: 부산
도착항: 첸나이
화물 준비일: 2026-11-20
조건: CIF
컨테이너: 40HQ 1대
인보이스 금액: USD 25,000
결제조건: T/T 30% Advance

감사합니다."""}],
     "expected": {"commodity": "전자부품", "qty": 300, "qtyUnit": "CTNS", "packing": "Carton", "totalCbm": 55,
                  "pol": "KRPUS", "pod": "INMAA", "cargoReadyDate": "2026-11-20", "incoterms": "CIF",
                  "containerType": "40HQ", "grossWeightKg": 9000, "invoiceValue": 25000, "ccy": "USD",
                  "paymentTerm": "T/T 30% Advance", "contactEmail": "sales@elec-parts.co.kr"}},
]


def all_cases() -> list[dict]:
    """기본 세트 + eval/cases/*.json (팀원·발주 측이 추가한 '블라인드' 메일).

    eval/cases/<이름>.json 형식:
      {"id": "T01", "desc": "설명", "mails": [{"file": "T01.eml"}, {"file": "T01_reply.eml", "role": "REPLY"}],
       "expected": {"commodity": "…", "qty": 10, …}, "multipleItems": false}
    file은 json과 같은 폴더 기준. expected에 없는 필드는 '메일에 없음'이 정답.
    """
    extra = []
    for path in sorted(CASES_DIR.glob("*.json")):
        case = json.loads(path.read_text(encoding="utf-8"))
        for mail in case["mails"]:
            mail["path"] = str(path.parent / mail.pop("file"))
        case.setdefault("desc", path.stem)
        extra.append(case)
    return CASES + extra
