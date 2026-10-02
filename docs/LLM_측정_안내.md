# 로컬 LLM 측정 안내 (GPU PC용)

B트랙의 남은 일은 **실제 LLM으로 추출 정확도를 재서(NFR-01) 쓸 모델을 정하는 것** 하나다.
평가 메일 24건(발주 측 샘플 4 + 변형 12 + 부산 기준 테스트 8)을 모델마다 돌려 한 표로 비교하고 추천까지 해 준다.

## 준비 (처음 한 번)
1. 레포 받기 → `server` 폴더에서 가상환경·패키지 설치 (설치 절차서 참고)
   ```
   cd server
   python -m venv .venv
   .venv\Scripts\pip install -r requirements.txt
   ```
2. https://ollama.com 에서 Ollama 설치 (이미 있으면 **최신 버전으로 업데이트** — Qwen3.5 같은 새 모델은 구버전에서 안 돌 수 있음)
3. `server\.env` 만들기 (메모장, 파일 형식 '모든 파일'로 저장해 `.env.txt`가 되지 않게):
   ```
   EXTRACTION_MODE=hybrid
   LLM_THINK=false
   LLM_SCOPE=all
   ```

## 측정 (더블클릭 한 번)
`server\run_benchmark.bat` 더블클릭 → 설정 점검 → 모델 받기 → 측정 → 보고서.
기본 모델은 `qwen3.5:9b`, `qwen3:14b` (RTX 4070 12GB 기준). 다른 모델은 cmd에서 `run_benchmark.bat qwen3:8b gemma4:12b`처럼.
노트북용 4070(8GB)이면 `run_benchmark.bat qwen3.5:9b`만.

모델 받기에 수 GB씩 내려받고, 측정은 모델당 몇 분 걸린다.

## 결과
`server\eval\reports\compare_날짜.md` — **이 파일을 공유**하면 된다.
- 맨 위 **결론**에 추천 모델이 적혀 있다 (정밀도 ≥90%, 메일 1건 ≤45초, LLM 오류 ≤10%를 통과한 모델 중 채움률 최고)
- **규칙만** 줄이 기준선. LLM이 채움률을 얼마나 올렸는지 본다. 특히 **T06(문장형)·V05·V06** 케이스
- 오추출 목록이 있으면 그대로 공유 — 프롬프트·검증 기준은 B트랙이 고친다 (팀원이 튜닝할 필요 없음)

## 정한 뒤
실제로 쓸 PC의 `server\.env`에 `LLM_MODEL=추천모델` 을 넣고 서버를 다시 켠다.
시연 PC가 GPU 없는 노트북이면 그 PC에서도 한 번 돌려 시간(45초)을 확인할 것 — 넘으면 `LLM_SCOPE=missing`.
