# 블라인드 평가 메일 넣는 곳

기본 평가 세트(`eval/dataset.py`)는 추출 룰을 만든 사람이 같이 만들어서 점수가 **낙관적**이다.
진짜 정밀도(NFR-01)는 룰을 보지 않은 사람이 쓴 메일로 재야 한다.

1. 견적 요청 메일을 `.eml`로 저장 (Outlook: 메일 열기 → 다른 이름으로 저장 / Gmail: 원본 보기 → 다운로드)
2. 같은 폴더에 정답 파일 `<이름>.json`:
```json
{"id": "T01", "desc": "무역회사 실제 문의 형식",
 "mails": [{"file": "T01.eml"}],
 "expected": {"commodity": "Plastic Parts", "qty": 500, "qtyUnit": "CTNS", "packing": "Carton",
              "pol": "KRPUS", "pod": "CNSHA", "grossWeightKg": 10000, "incoterms": "FOB",
              "contactEmail": "buyer@example.com"},
 "multipleItems": false}
```
- 값은 6장 형식(kg·CBM·mm·UN/LOCODE·YYYY-MM-DD). **메일에 없는 필드는 적지 않는다** (없음이 정답).
- 회신이 있으면 `{"file": "T01_reply.eml", "role": "REPLY"}` 를 뒤에 추가.
3. `python -m eval.run_eval` — 자동으로 포함된다.

실제 고객 메일은 개인정보·외부 재배포 금지 규칙을 지킬 것 (저장소가 비공개인지 확인).
