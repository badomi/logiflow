"""B트랙 추출 구성요소 단위 테스트 — 정의서 수용 기준별로 묶었다."""

import io

import pytest
from openpyxl import Workbook

from app.extraction import attachments, engine, grounding, parsers as P, rules, signature
from app.extraction.base import MailText, Source
from app.extraction.preprocess import split_mail
from app.llm import LlmError, ensure_local


def run_rules(body: str, subject: str = "견적 요청", sender=("", "a@b.co")) -> dict:
    return engine.run([MailText(subject, sender[0], sender[1], body)])["fields"]


# ---------------------------------------------------------------- FR-203 단위·코드 정규화 (결정형)


@pytest.mark.parametrize("text,kg", [("2,650 LBS", 1202.02), ("12.5 MT", 12500), ("약 1,419 kg", 1419), ("3 톤", 3000)])
def test_weight_to_kg(text, kg):
    assert P.parse_weight_kg(text).value == kg


@pytest.mark.parametrize("text,cbm", [("350 CFT", 9.9109), ("28 m3", 28), ("0.19CBM", 0.19)])
def test_volume_to_cbm(text, cbm):
    assert P.parse_cbm(text).value == cbm


@pytest.mark.parametrize("text,mm", [("31x45x27 cm", (310, 450, 270)), ('12" x 18" x 10"', (305, 457, 254)),
                                     ("12 x 18 x 10 inch", (305, 457, 254)), ("L310 W450 H270 mm", (310, 450, 270))])
def test_dimensions_to_mm(text, mm):
    assert P.parse_dims_mm(text).value == mm


def test_ambiguous_values_are_not_guessed():
    assert P.parse_date("11월 5일") is None  # 연도 없음
    assert P.parse_dims_mm("680 x 400 x 370").conf < 0.75  # 단위 없음 → 검토 필요 수준
    assert P.parse_qty("33,000개").conf < 0.75  # 낱개 수량 ≠ 포장 수량
    assert P.port_code("부산 또는 인천") is None  # 항구 둘
    assert P.parse_container("20개") is None


def test_ports_and_candidates():
    assert P.port_code("BUSAN, KOREA") == "KRPUS" and P.port_code("호주 시드니") == "AUSYD"
    assert P.port_code("Los Angeles") == "USLAX" and P.port_code("KRINC") == "KRINC"
    assert P.port_code("SIDNEY") is None and "AUSYD" in P.port_candidates("SIDNEY")


# ---------------------------------------------------------------- FR-205 국·영문 혼용


def test_korean_and_english_labels_give_same_result():
    ko = run_rules("품목: 플라스틱 부품\n수량: 500박스\n총중량: 10,000kg\n선적항: 부산\n도착항: 상하이\n조건: FOB")
    en = run_rules("Commodity: Plastic Parts\nQ'ty: 500 CTNS\nG.W: 10,000 KGS\nPOL: Busan\nPOD: Shanghai\nIncoterms: FOB")
    for key in ("qty", "qtyUnit", "grossWeightKg", "pol", "pod", "incoterms"):
        assert ko[key]["value"] == en[key]["value"] and en[key]["status"] == "filled", key


def test_label_similarity_tolerates_variants():
    assert rules.label_scores("Gross Weight(KG)")[0][0] == "grossWeightKg"
    assert rules.label_scores("박스 1개 규격")[0][0] == "box"
    assert rules.label_scores("Port of Loading")[0][0] == "pol"


# ---------------------------------------------------------------- 전처리: 인용문(WBS 3.5)·목록·서명


def test_quote_signature_and_broken_list():
    body = ("회신드립니다.\r\n  1.\r\n화물 준비일: 2026-10-20\r\n감사합니다.\r\n홍길동 과장\r\n㈜에이비씨무역\r\n"
            "________________\r\n보낸 사람: LEONA\r\n품목: 이전 메일 값")
    parts = split_mail(body)
    assert "1. 화물 준비일: 2026-10-20" in parts.body
    assert "이전 메일" not in parts.body and "이전 메일" in parts.quoted
    assert parts.signature == "홍길동 과장\n㈜에이비씨무역"


def test_gmail_quote_lines_removed():
    parts = split_mail("준비일: 2026-10-20\n\n2026년 9월 29일 (화) 오전 11:46, 박재석 <a@b.c>님이 작성:\n> 품목: X")
    assert parts.body == "준비일: 2026-10-20"


# ---------------------------------------------------------------- FR-208 고객사·담당자·회신주소


def test_signature_people():
    ko = {c.field: c.value for c in signature.extract("김민수 대리\n㈜한빛무역 | 해외영업팀\nTel. 051-123-4567", "", "ms@hanbit.co.kr", "메일 1", 1)}
    assert ko == {"contactEmail": "ms@hanbit.co.kr", "contactName": "김민수 대리", "customerName": "㈜한빛무역"}
    en = {c.field: c.value for c in signature.extract("John Smith\nSales Manager\nABC Trading Co., Ltd.", "John Smith", "js@abc.com", "메일 1", 1)}
    assert en["contactName"] == "John Smith" and en["customerName"] == "ABC Trading Co., Ltd."


def test_generic_display_name_is_not_a_person():
    fields = run_rules("품목: 노트\n감사합니다.", sender=("화주 담당자", "shipper@example.com"))
    assert fields["contactName"]["status"] == "missing" and fields["contactEmail"]["value"] == "shipper@example.com"


# ---------------------------------------------------------------- FR-202·209 근거·점수·임계치


def test_every_field_has_evidence_and_score():
    fields = run_rules("품목: 노트\n박스 규격: 680 x 400 x 370\n총중량: 500kg")
    assert fields["commodity"] == {**fields["commodity"], "value": "노트", "evidence": "품목: 노트", "status": "filled"}
    assert fields["boxL"]["status"] == "review" and fields["boxL"]["value"] == 680  # 단위 없는 규격 → 검토 필요
    assert fields["pol"]["status"] == "missing" and fields["pol"]["value"] is None


def test_grounding_rejects_quotes_not_in_mail():
    src = [Source(101, "메일 1 · 본문", "body", "품목: 노트\n총중량: 500kg")]
    assert grounding.ground("총중량: 500kg", "500kg", src)[0] == 1.0
    assert grounding.ground("총중량: 5,000kg", "5,000kg", src)[0] == 0.0  # 원문과 다른 숫자
    assert grounding.ground("결제: T/T", "T/T", src)[0] == 0.0  # 원문에 없는 문장
    assert grounding.ground("총중량: 500kg", "700kg", src)[0] == 0.0  # 인용은 맞지만 값이 인용에 없음


def test_newer_reply_overrides_older_value():
    result = engine.run([
        MailText("요청", "", "a@b.co", "총중량: 500kg"),
        MailText("RE: 요청", "", "a@b.co", "총중량 정정합니다.\n총중량: 650kg", role="REPLY"),
    ])
    assert result["fields"]["grossWeightKg"]["value"] == 650 and "메일 2" in result["fields"]["grossWeightKg"]["origin"]


# ---------------------------------------------------------------- FR-210 다품목


def test_multi_item_sections_detected():
    body = "[화장품]\n품목: 화장품\n총 박스 수: 330 CTN\n[의류]\n품목: 의류\n총 박스 수: 1,000 CTN"
    result = engine.run([MailText("견적", "", "a@b.co", body)])
    assert result["multipleItems"] and result["fields"]["qty"]["status"] == "review"


# ---------------------------------------------------------------- FR-204 첨부


def test_xlsx_packing_list_text_is_extracted():
    wb = Workbook()
    ws = wb.active
    ws.append(["Gross Weight", "1,200 KGS"])
    ws.append(["Total CBM", "3.0 CBM"])
    buf = io.BytesIO()
    wb.save(buf)
    text = attachments.extract_text("packing.xlsx", buf.getvalue())
    result = engine.run([MailText("견적", "", "a@b.co", "첨부 참고 부탁드립니다.", attachments=[("packing.xlsx", text)])])
    assert result["fields"]["grossWeightKg"]["value"] == 1200
    assert result["fields"]["totalCbm"]["origin"] == "메일 1 · 첨부 packing.xlsx"


# ---------------------------------------------------------------- NFR-04


def test_llm_must_be_local():
    ensure_local("http://localhost:11434")
    ensure_local("http://192.168.0.10:11434")
    with pytest.raises(LlmError, match="NFR-04"):
        ensure_local("http://8.8.8.8:11434")



def test_llm_prompt_and_schema_only_contain_asked_items():
    from app.extraction import llm_stage

    src = [Source(101, "메일 1 · 본문", "body", "노트 10박스"), Source(150, "메일 1 · 서명", "signature", "김민수 대리")]
    prompt = llm_stage.build_prompt(src, ["commodity", "quantity"])
    assert "- commodity" in prompt and "- grossWeight" not in prompt
    assert "김민수" not in prompt  # 고객 정보를 묻지 않으면 서명도 보내지 않는다 (입력 줄이기)
    assert set(llm_stage.schema_for(["commodity"])["properties"]["fields"]["properties"]) == {"commodity"}
