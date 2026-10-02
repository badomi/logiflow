"""docs/samples.md의 샘플 메일로 테스트용 .eml 파일을 만든다.

발주 측 샘플은 교육용 가공 자료이며 외부 재배포 금지다. 만들어진 파일도 같은 취급을 한다.
첨부파일(인보이스 등)은 Notion에 실물이 없어 최소한의 가짜 PDF로 대신한다.

실행 (server 폴더에서): .venv\\Scripts\\python scripts\\make_sample_emls.py
"""

from email.message import EmailMessage
from email.utils import format_datetime
from datetime import datetime, timedelta, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "samples" / "eml"
KST = timezone(timedelta(hours=9))

SHIPPER = ("화주 담당자", "shipper@example.com")  # 견적 요청하는 고객(화주) 역할
LEONA = ("LEONA 견적 담당자", "quote@leona.example.com")  # LEONA 담당자 역할

LCL_BODY = """아래 조건으로 해상 LCL 운임 견적을 요청드립니다.
품목: 스프링노트(일반 문구류)
픽업지: 인천
도착항: 호주 시드니
박스 수량: 10박스
박스 1개 규격: 310 × 450 × 270mm
박스당 부피: 약 0.038CBM
전체 부피: 약 0.19CBM
예상 총중량: 최대 500kg
조건: FOB 및 CIF 산출용
선적항은 별도로 정하지 않았으므로, 남양주 픽업 기준으로 비용과 운송 일정이 유리한 항구를 적용해 주세요.
아래 항목을 포함한 금액으로 안내 부탁드립니다.
공장 → 선적항 내륙운송비
수출통관비
THC 및 서류비
필요시 CFS 창고료
시드니까지의 해상운임
적하보험료 또는 보험료 적용 기준
1CBM 미만 화물에 최소 청구 운임이 적용되는 경우 해당 금액도 함께 안내 부탁드립니다.
포함되지 않은 비용이 있다면 별도로 표시해 주시고, 견적 유효기간과 예상 운송기간도 함께 부탁드립니다.
감사합니다.
"""

FOB_BODY = """안녕하세요.
수출 건 관련하여 FOB 조건의 국내 운송 및 선적비용 견적 문의드립니다.
[화장품]
출고지: 부산광역시 사하구
선적항: 부산항
품목: 화장품
제품 중량: 150g/EA
수량: 33,000개
포장단위: 100개/CTN
총 박스 수: 330 CTN
박스 사이즈: 680 × 400 × 370 mm
총 부피: 약 33.21 CBM
총 Gross Weight: 약 1,419 kg
Pallet Size: 1,100 × 1,100 mm
팔레트 적재수량: 20 CTN/PLT
예상 Pallet 수: 약 17 PLT
보관조건: 상온
운송조건: FOB
선적방식: FCL
출고 가능일: 미정
[의류]
출고지: 대전광역시 서구
선적항: 출고지 기준 적합한 선적항으로 제안 부탁드립니다.
품목: 의류
제품 용량: 300g/EA
수량: 30,000개
포장단위: 30개/CTN
총 박스 수: 1,000 CTN
박스 사이즈: 460 × 285 × 245 mm
총 부피: 약 32.12 CBM
총 Gross Weight: 약 4,950 kg
보관조건: 상온
운송조건: FOB
선적방식: FCL
제조사 기준 예상 컨테이너: 40FT HC 1대
출고 가능일: 미정
FOB 조건으로 공장 출고지부터 본선 적재까지 발생하는 국내 운송비 및 수출 제비용 견적 부탁드립니다.
상기 물량 기준으로 적합한 컨테이너 규격 및 필요 대수도 함께 확인 부탁드립니다.
견적서에는 내륙운송비, 터미널 및 선적 관련 비용, 수출통관 및 서류 관련 비용 등 국내에서 발생하는 비용을 가능한 한 항목별로 구분하여 안내 부탁드립니다.
또한 견적에 포함되지 않고 실제 선적 진행 시 추가로 발생할 수 있는 비용이 있다면 해당 항목도 함께 안내 부탁드립니다.
추가로 견적 산출을 위해 필요한 정보나 확인사항이 있다면 말씀 부탁드립니다.
"""

INSURANCE_BODY = """안녕하세요, 업무에 노고 많으십니다.
유첨의 건 보험 부보 요청드립니다!

· 실화주 : G화주
· POL/POD : BUSAN – GENOA
· ETD/ETA : 9/8 – 11/1
· VSSL : SAMPLE OCEAN 0002W

감사합니다.
"""

SPACE_HTML = """<html><body>
<p>안녕하세요, 담당자님 업무에 노고 많으십니다.<br>하기 모선에 40HQ*1 추가 가능할지 스페이스 확인해 주시면 감사하겠습니다!</p>
<table border="1">
<tr><th>항목</th><th>내용</th></tr>
<tr><td>선명/항차</td><td>SAMPLE STAR 0001S (ROUTE : XXX)</td></tr>
<tr><td>출발</td><td>BUSAN, KOREA / 입항 09.11 09:30 &middot; 출항 09.12 06:30 / 부산항 X터미널</td></tr>
<tr><td>도착</td><td>KEELUNG, TAIWAN / 입항 09.14 19:30 / 현지 Y터미널</td></tr>
<tr><td>소요시간</td><td>2일 13시간</td></tr>
<tr><td>Container 반입마감</td><td>09.11 03:30</td></tr>
<tr><td>서류마감</td><td>09.09 11:00</td></tr>
<tr><td>요청 수량</td><td>40HQ &times; 1</td></tr>
</table>
<style>p { color: black; }</style>
</body></html>"""


def fake_pdf(title: str) -> bytes:
    """글자 한 줄짜리 최소 PDF (첨부 처리 테스트용)."""
    stream = f"BT /F1 18 Tf 72 720 Td ({title}) Tj ET".encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, xref)
    return bytes(out)


def message(subject: str, sender, recipient, sent: datetime, msgid: str, **extra) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f'"{sender[0]}" <{sender[1]}>'
    msg["To"] = f'"{recipient[0]}" <{recipient[1]}>'
    msg["Date"] = format_datetime(sent)
    msg["Message-ID"] = msgid
    for name, value in extra.items():
        msg[name.replace("_", "-")] = value
    return msg


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    base = datetime(2026, 9, 28, 9, 0, tzinfo=KST)

    # 01: 실제 입력 형태 — LCL 견적 요청 (불일치 포함: 인천 vs 남양주, 0.038×10 ≠ 0.19)
    lcl_id = "<sample-01-lcl@example.com>"
    m = message("해상 LCL 운임 견적 요청 (인천 → 시드니)", SHIPPER, LEONA, base, lcl_id)
    m.set_content(LCL_BODY)
    (OUT / "01_lcl_request.eml").write_bytes(bytes(m))

    # 02: 실제 입력 형태 — FOB 다품목 (FR-210 수동 처리 대상)
    m = message("FOB 조건 국내 운송 및 선적비용 견적 문의", SHIPPER, LEONA, base + timedelta(hours=1),
                "<sample-02-fob@example.com>")
    m.set_content(FOB_BODY)
    (OUT / "02_fob_multi_item.eml").write_bytes(bytes(m))

    # 03: 01에 대한 화주 회신 — 헤더(In-Reply-To/References)로 회신 식별 (FR-104)
    m = message("RE: 해상 LCL 운임 견적 요청 (인천 → 시드니)", SHIPPER, LEONA, base + timedelta(days=1),
                "<sample-03-lcl-reply@example.com>", In_Reply_To=lcl_id, References=lcl_id)
    m.set_content("요청하신 정보 회신드립니다.\n화물 준비일: 2026-10-20\n인보이스 금액: USD 3,000\n\n"
                  "> 아래 조건으로 해상 LCL 운임 견적을 요청드립니다.\n> 품목: 스프링노트(일반 문구류)\n")
    (OUT / "03_lcl_reply.eml").write_bytes(bytes(m))

    # 04: 첨부파일 2개 (인보이스·패킹리스트) — 원문·첨부 분리 저장 (FR-103)
    m = message("[C보험사] G화주 보험 부보 요청의 건 / REF-0001", ("수출 담당자", "export@a-forwarding.example.com"),
                ("보험사 담당자", "cargo@c-insurance.example.com"), base + timedelta(hours=2),
                "<sample-04-insurance@example.com>")
    m.set_content(INSURANCE_BODY)
    m.add_attachment(fake_pdf("COMMERCIAL INVOICE (SAMPLE)"), maintype="application", subtype="pdf",
                     filename="인보이스_REF-0001.pdf")
    m.add_attachment(fake_pdf("PACKING LIST (SAMPLE)"), maintype="application", subtype="pdf",
                     filename="패킹리스트_REF-0001.pdf")
    (OUT / "04_insurance_with_attachments.eml").write_bytes(bytes(m))

    # 05: HTML 본문 + 표 — 견적 요청이 아닌 업무 메일, HTML→텍스트 변환 확인
    m = message("[B선사/A포워딩] BUSAN - HONG KONG 스페이스 문의의 건", ("수출 담당자", "export@a-forwarding.example.com"),
                ("담당자", "space@b-line.example.com"), base + timedelta(hours=3),
                "<sample-05-space@example.com>")
    m.set_content(SPACE_HTML, subtype="html")
    (OUT / "05_space_inquiry_html.eml").write_bytes(bytes(m))

    # 06: 오래된 한국어 메일 문자셋(EUC-KR) — 문자 깨짐 없이 읽는지 확인
    m = message("해상 LCL 운임 견적 요청 (EUC-KR 인코딩)", SHIPPER, LEONA, base + timedelta(hours=4),
                "<sample-06-euckr@example.com>")
    m.set_content(LCL_BODY, charset="euc-kr")
    (OUT / "06_lcl_request_euckr.eml").write_bytes(bytes(m))

    for file in sorted(OUT.glob("*.eml")):
        print(file.relative_to(OUT.parent.parent))


if __name__ == "__main__":
    main()
