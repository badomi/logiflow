""".eml 파서: 애드인 원문과 .eml 리더가 같은 결과를 내야 한다."""

from app.mail_parser import html_to_text, parse_eml

from .conftest import sample


def test_plain_korean_mail():
    snap = parse_eml(sample("01_lcl_request.eml")).snapshot
    assert snap.subject == "해상 LCL 운임 견적 요청 (인천 → 시드니)"
    assert snap.from_.email == "shipper@example.com"
    assert snap.to[0].email == "quote@leona.example.com"
    assert snap.internet_message_id == "<sample-01-lcl@example.com>"
    assert "박스당 부피: 약 0.038CBM" in snap.body_text
    assert snap.received_at is not None


def test_reply_headers():
    snap = parse_eml(sample("03_lcl_reply.eml")).snapshot
    assert snap.in_reply_to == "<sample-01-lcl@example.com>"
    assert snap.references == ["<sample-01-lcl@example.com>"]


def test_attachments_are_separated():
    parsed = parse_eml(sample("04_insurance_with_attachments.eml"))
    names = [a.filename for a in parsed.attachments]
    assert names == ["인보이스_REF-0001.pdf", "패킹리스트_REF-0001.pdf"]
    assert all(a.data.startswith(b"%PDF") for a in parsed.attachments)
    assert [a.filename for a in parsed.snapshot.attachments] == names
    assert all(len(a.sha256) == 64 for a in parsed.snapshot.attachments)


def test_html_body_becomes_text_with_table_cells():
    body = parse_eml(sample("05_space_inquiry_html.eml")).snapshot.body_text
    assert "<" not in body and "p {" not in body  # 태그·스타일 제거
    assert "선명/항차\tSAMPLE STAR 0001S" in body  # 표 칸 구분 유지
    assert "KEELUNG, TAIWAN" in body


def test_euc_kr_mail_is_not_broken():
    snap = parse_eml(sample("06_lcl_request_euckr.eml")).snapshot
    assert "스프링노트" in snap.body_text


def test_html_to_text_decodes_entities():
    assert html_to_text("<p>40HQ &times; 1&nbsp;&nbsp;대</p>") == "40HQ × 1 대"
