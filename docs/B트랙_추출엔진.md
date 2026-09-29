# B트랙 추출 엔진 — 룰 + 로컬 LLM 하이브리드

코드: `server/app/extraction/` · 측정: `server/eval/` · 테스트: `server/tests/test_extraction.py`, `test_pipeline_flow.py`, `test_eval_precision.py`

## 설계 원칙
**LLM은 "어디에 무슨 값이 적혀 있는지" 찾아 원문 그대로 옮기기만 한다. 변환·판단은 전부 룰.**

```
메일(원문+회신+첨부)
  └ preprocess   인용문 분리(WBS 3.5) · Outlook 목록 갈라짐 복구 · 서명 분리
      ├ rules        키워드 사전(국·영문) '라벨: 값' + 유사도 → 후보(점수)          FR-201·205
      ├ signature    서명·From 헤더 → 고객사·담당자·이메일                          FR-208
      └ llm_stage    로컬 LLM: {value: 원문 그대로, quote: 근거 문장 그대로}
           └ grounding  인용이 원문에 있나? 숫자가 같나? 값이 인용 안에 있나? → 아니면 버림 (환각 차단)
  └ parsers      LBS→KG, CFT→CBM, inch/cm→mm, 날짜 ISO, 항구→UN/LOCODE (결정형)     FR-203
  └ engine       필드별 결합: 룰=LLM 일치 +0.1 / 불일치 → 검토 필요 / 최신 메일 우선    FR-202·207·209
                 임계치(fields.json) 이상만 저장. 다품목이면 화물 필드 저장 안 함        FR-210
```
LLM이 꺼져 있거나 느리면(45초 초과) **룰 결과로 계속** 진행한다(NFR-03). 외부 주소 LLM은 실행 거부(NFR-04).

**LLM에게는 룰이 못 찾은 항목만 묻는다** (CPU만 있는 노트북에서 60초 안에 끝내려고. 측정: 내장그래픽 qwen3:4b 출력 13토큰/초).
- 룰이 전부 찾았거나, 비어 있는 게 고객사명·담당자명뿐이면 LLM을 부르지 않는다 (`llm.skipped`)
- 룰이 찾은 값이라도 **최신 회신이 그 항목을 문장으로 언급**하면 다시 묻는다 ("Gross weight is revised to 13,000 KGS")
- 다품목·조건 여러 개(FOB 및 CIF)는 LLM도 못 정하므로 묻지 않는다
- 결과 JSON의 `llm.asked`에 물어본 항목이 남는다

## 켜는 법 (Windows)
1. https://ollama.com 설치 → PowerShell: `ollama pull qwen3:8b`
2. `server\.env`: `EXTRACTION_MODE=hybrid` (LLM 없이: `rules`, 끄기: `off`)
3. 견적까지 보려면 테스트 요율 등록: `.venv\Scripts\python -m app.seed --sample-rates` (다시 실행해도 중복 안 됨)
   - 출처: LCL·FOB 예시 견적서(부대비용) + `samples.md` 3번 D선사 월간 운임표(부산 출발 17개 항구 해상운임, POL 부산 가정)
   - 견적이 나오는 조합: **인천→시드니 LCL**, **부산 40HQ FOB(도착항 무관)**, **부산→월간 운임표 17개 항구 40HQ CIF**
   - 20ft·냉동 컨테이너는 해상운임은 있지만 국내 부대비용 자료가 없어 "국내 부대비용 요율 없음"으로 보류 (불완전 견적 금지)
   - `samples.md` 4번 미주 운임은 미등록: 컨테이너 크기 미표기, 같은 구간 복수 선사, 철도 터미널(내륙) 설계 필요

## 정밀도 측정 (NFR-01) — B트랙 핵심 산출물
```powershell
cd server
.\.venv\Scripts\python -m eval.run_eval                  # 룰만
.\.venv\Scripts\python -m eval.run_eval --mode hybrid    # 룰 + LLM
.\.venv\Scripts\python -m eval.run_eval --mode hybrid --model qwen2.5:14b-instruct   # 모델 비교
```
보고서는 `server/eval/reports/`. **오추출 목록**이 W3–4 "오추출 사례 피드백" 자료.

| 측정 (2026-09-29, 룰만) | 결과 |
|---|---|
| 정밀도 | 99.5% (206/207) |
| 채움률 | 92.0% (206/224) |
| 처리 시간 | 최대 0.05초 |

⚠ **이 숫자는 낙관적이다.** 평가 메일(`eval/dataset.py`)을 룰과 같은 사람이 만들었다.
실제 정밀도는 룰을 안 본 사람이 쓴 메일을 `eval/cases/`에 넣고 재야 한다 (방법: `eval/cases/README.md`).
남은 오추출 1건(V06: 회신 문장 "Gross weight is revised to 13,000 KGS")과 문장형 메일(V05, 채움 7/16)이 LLM이 메울 곳.

## 모델 고르기
- 기본값 `qwen3:8b` — 다국어, **Apache 2.0** (Qwen3 전 크기). GPU 12GB 이상이면 `qwen3:14b`, GPU 없으면 `qwen3:4b`도 측정.
- Qwen3는 기본이 생각(thinking) 모드라 느리다 → `LLM_THINK=false`(기본값)로 끈다. 생각 모드가 없는 모델은 자동으로 옵션 없이 요청.
- 후보가 여럿이면 `run_eval --mode hybrid --model …`로 **정밀도·시간을 재서** 고른다. 감으로 고르지 않는다.
- **GPU 없는 노트북은 60초(NFR-02)를 넘길 수 있다.** 넘으면 룰 결과로 진행되지만(케이스는 멈추지 않음) LLM 효과가 없다.
  보고서의 케이스별 시간·"LLM 오류: timed out"을 확인할 것. 발주 측 PC 사양 확인 필요.

## 데이터로 관리하는 것 (NFR-06, 코드 수정 없음)
- `app/extraction/data/fields.json` — 필드별 국·영문 라벨 사전, 필드별 임계치(FR-209), 다품목 섹션 패턴
- `app/extraction/data/ports.json` — 항구명 → UN/LOCODE (사내 항구 마스터 받으면 교체)
- 고친 뒤 `pytest` → `test_eval_precision.py`가 정밀도 하락을 잡는다

## DB (`server/app/models.py`)
- `quote_inputs` — 6장 필드 (정의서 이름·타입 그대로). **filled 값만** 저장
- `extraction_fields` — 추출할 때마다 21개 필드 전부: 값·근거 원문·점수·방법·상태 (FR-202, NFR-05). `method=manual`은 담당자 수정(FR-206) → 재추출이 덮어쓰지 않음

## API (D트랙 콘솔·A트랙 패널용)
- `GET /api/cases/{id}/fields` — 필드별 값·근거·점수·상태(filled/review/missing), 검토 필요 값은 `candidate`
- `POST /api/cases/{id}/fields/{field}` `{"value": …, "actor": …}` — 수동 보정 → 이력 → 검증 재실행
- `POST /api/cases/{id}/extract` `{"actor": …}` — 다시 추출 (NFR-03). 패널 '항목 추출' 영역의 [다시 추출] 버튼.
  설정을 바꾼 뒤, LLM 시간 초과로 규칙 결과만 나왔을 때 쓴다. 담당자가 고친 필드는 유지. 실행 중이면 409 (2분 넘게 멈춘 경우는 허용)

## A트랙 임시 부분을 추출·검증 결과로 연결함 (2026-09-29)
- 패널 '항목 추출' 영역: `result.fields` 형식을 읽어 값 표시, 확신 낮은 값은 "(확인 필요)", 마우스를 올리면 근거 원문
- 패널 '보완 요청 문항': 검증 결과 문항을 자동으로 채움 (담당자가 고치면 그 내용 유지, 다른 메일로 가면 초기화)
- 보완 요청 초안: 문항을 비워 보내면 검증 결과 문항 사용 (FR-305)
- 초안 받는 사람: 추출한 `contactEmail`·`contactName` 우선 (FR-208)
- `CLAUDE.md`·`docs/트랙간_전달사항.md`의 "B·C 연결 전 임시" 기록은 A트랙 문서라 고치지 않았다 → A트랙이 갱신

## 다른 트랙에 영향
- **C트랙**: 추출 결과 JSON의 `fields[*].status == "review"`는 "확인 필요" 문항으로 물으면 된다(`validation.py`에 예시 구현).
  `notes`의 `MULTIPLE_INCOTERMS`·`PORT_UNMAPPED`(후보 포함)도 참고.
- **A트랙**: 인용문 분리는 B가 하므로 원문을 계속 자르지 않고 넘기면 된다. 초안 받는 사람은 `contactEmail`로 바꿔도 된다(FR-208).
- **D트랙**: 위 두 API로 필드 화면(근거·점수 표시, 수정)을 만들 수 있다.

## 아직 안 된 것
- 실제 LLM으로 측정 안 함(개발 환경에서 Ollama 실행 불가) → 팀 PC에서 `--mode hybrid` 측정 필요
- 블라인드 평가 메일 0건, MUST-SHIP ① 시연용 샘플 20건 중 16건
- 스캔 PDF OCR(FR-204 선택 사항) 없음
